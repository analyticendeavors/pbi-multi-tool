"""
Schema Validator - Compare source and target model structures
Built by Reid Havens of Analytic Endeavors

Validates that target model has compatible structure before swapping.
This is a warn-only validation - shows differences but allows proceeding.
"""

import logging
from dataclasses import dataclass, field
from typing import List, Optional, Set, Dict, Any
from enum import Enum

from tools.connection_hotswap.models import ConnectionMapping, SwapTarget


class ValidationSeverity(Enum):
    """Severity level of validation findings."""
    INFO = "info"           # Informational, not a problem
    WARNING = "warning"     # Potential issue, may cause problems
    ERROR = "error"         # Serious issue, likely to cause problems


@dataclass
class ValidationFinding:
    """A single validation finding."""
    severity: ValidationSeverity
    category: str           # e.g., "Table", "Column", "Measure"
    message: str
    details: str = ""

    @property
    def icon(self) -> str:
        """Get icon for this severity level."""
        return {
            ValidationSeverity.INFO: "ℹ️",
            ValidationSeverity.WARNING: "⚠️",
            ValidationSeverity.ERROR: "❌",
        }.get(self.severity, "•")


@dataclass
class SchemaInfo:
    """Schema information extracted from a model."""
    tables: Set[str] = field(default_factory=set)
    columns: Dict[str, Set[str]] = field(default_factory=dict)  # table -> columns
    measures: Set[str] = field(default_factory=set)
    relationships: int = 0
    calculated_tables: Set[str] = field(default_factory=set)


@dataclass
class ValidationResult:
    """Result of schema validation."""
    source_name: str
    target_name: str
    findings: List[ValidationFinding] = field(default_factory=list)
    source_schema: Optional[SchemaInfo] = None
    target_schema: Optional[SchemaInfo] = None

    @property
    def has_errors(self) -> bool:
        """Check if there are any error-level findings."""
        return any(f.severity == ValidationSeverity.ERROR for f in self.findings)

    @property
    def has_warnings(self) -> bool:
        """Check if there are any warning-level findings."""
        return any(f.severity == ValidationSeverity.WARNING for f in self.findings)

    @property
    def is_compatible(self) -> bool:
        """Check if schemas appear to be compatible (no errors)."""
        return not self.has_errors

    @property
    def summary(self) -> str:
        """Get a summary of the validation result."""
        errors = sum(1 for f in self.findings if f.severity == ValidationSeverity.ERROR)
        warnings = sum(1 for f in self.findings if f.severity == ValidationSeverity.WARNING)
        infos = sum(1 for f in self.findings if f.severity == ValidationSeverity.INFO)

        parts = []
        if errors > 0:
            parts.append(f"{errors} error(s)")
        if warnings > 0:
            parts.append(f"{warnings} warning(s)")
        if infos > 0:
            parts.append(f"{infos} info(s)")

        if not parts:
            return "✅ Schemas appear compatible"
        return ", ".join(parts)


class SchemaValidator:
    """
    Validates that source and target models have compatible structures.

    Note: This performs a structural comparison only. It cannot verify
    that data types or actual data are compatible.
    """

    def __init__(self, connector):
        """
        Initialize the schema validator.

        Args:
            connector: PowerBIConnector instance with active connection
        """
        self.logger = logging.getLogger(__name__)
        self.connector = connector

    def validate_mapping(self, mapping: ConnectionMapping) -> ValidationResult:
        """
        Validate schema compatibility between source and target.

        Args:
            mapping: ConnectionMapping with source and target

        Returns:
            ValidationResult with findings
        """
        if not mapping.target:
            return ValidationResult(
                source_name=mapping.source.display_name,
                target_name="(no target)",
                findings=[ValidationFinding(
                    severity=ValidationSeverity.ERROR,
                    category="Configuration",
                    message="No target configured",
                    details="Configure a target before validation"
                )]
            )

        result = ValidationResult(
            source_name=mapping.source.display_name,
            target_name=mapping.target.display_name
        )

        try:
            # Get source schema from current model
            source_schema = self._extract_source_schema()
            result.source_schema = source_schema

            # Get target schema (if accessible)
            target_schema = self._extract_target_schema(mapping.target)
            result.target_schema = target_schema

            if source_schema and target_schema:
                # Compare schemas
                self._compare_schemas(source_schema, target_schema, result)
            elif not target_schema:
                result.findings.append(ValidationFinding(
                    severity=ValidationSeverity.WARNING,
                    category="Connection",
                    message="Cannot access target schema",
                    details="Target model may not be running or accessible. Validation skipped."
                ))

        except Exception as e:
            self.logger.error(f"Schema validation error: {e}")
            result.findings.append(ValidationFinding(
                severity=ValidationSeverity.ERROR,
                category="Validation",
                message=f"Validation error: {str(e)}"
            ))

        return result

    def _extract_source_schema(self) -> Optional[SchemaInfo]:
        """Extract schema from the currently connected source model."""
        try:
            if not self.connector._model:
                return None

            schema = SchemaInfo()
            model = self.connector._model

            # Extract tables
            for table in model.Tables:
                table_name = str(table.Name)
                schema.tables.add(table_name)
                schema.columns[table_name] = set()

                # Check if calculated table
                if table.Partitions and len(table.Partitions) > 0:
                    partition = table.Partitions[0]
                    source_type = str(partition.SourceType) if hasattr(partition, 'SourceType') else ""
                    if 'Calculated' in source_type:
                        schema.calculated_tables.add(table_name)

                # Extract columns
                for column in table.Columns:
                    schema.columns[table_name].add(str(column.Name))

                # Extract measures
                for measure in table.Measures:
                    schema.measures.add(str(measure.Name))

            # Count relationships
            if hasattr(model, 'Relationships'):
                schema.relationships = len(model.Relationships)

            return schema

        except Exception as e:
            self.logger.error(f"Error extracting source schema: {e}")
            return None

    def _extract_target_schema(self, target: SwapTarget) -> Optional[SchemaInfo]:
        """
        Extract schema from a target model.

        Note: This requires connecting to the target, which may not always
        be possible. For local targets, we try to connect. For cloud targets,
        we skip (would require separate auth).
        """
        try:
            if target.target_type == "local":
                return self._extract_local_target_schema(target)
            else:
                # Cloud targets require XMLA auth - skip detailed validation
                self.logger.info("Cloud target - skipping detailed schema extraction")
                return None

        except Exception as e:
            self.logger.error(f"Error extracting target schema: {e}")
            return None

    def _extract_local_target_schema(self, target: SwapTarget) -> Optional[SchemaInfo]:
        """Extract schema from a local Power BI Desktop model."""
        try:
            # Import here to avoid circular dependency
            import clr
            from System import String

            # Create a temporary connection to the target
            # Note: This uses a separate connection, not the main connector
            try:
                TOM = clr.GetClrType("Microsoft.AnalysisServices.Tabular")
            except:
                # TOM assembly already loaded through connector
                pass

            from Microsoft.AnalysisServices.Tabular import Server

            server = Server()
            connection_string = f"Provider=MSOLAP;Data Source={target.server}"

            try:
                server.Connect(connection_string)

                if not server.Databases or len(server.Databases) == 0:
                    return None

                # Find the target database
                target_db = None
                for db in server.Databases:
                    if str(db.Name) == target.database or str(db.ID) == target.database:
                        target_db = db
                        break

                if not target_db:
                    # Use first database if exact match not found
                    target_db = server.Databases[0]

                schema = SchemaInfo()
                model = target_db.Model

                # Extract tables
                for table in model.Tables:
                    table_name = str(table.Name)
                    schema.tables.add(table_name)
                    schema.columns[table_name] = set()

                    # Check if calculated table
                    if table.Partitions and len(table.Partitions) > 0:
                        partition = table.Partitions[0]
                        source_type = str(partition.SourceType) if hasattr(partition, 'SourceType') else ""
                        if 'Calculated' in source_type:
                            schema.calculated_tables.add(table_name)

                    # Extract columns
                    for column in table.Columns:
                        schema.columns[table_name].add(str(column.Name))

                    # Extract measures
                    for measure in table.Measures:
                        schema.measures.add(str(measure.Name))

                # Count relationships
                if hasattr(model, 'Relationships'):
                    schema.relationships = len(model.Relationships)

                return schema

            finally:
                try:
                    server.Disconnect()
                except:
                    pass

        except Exception as e:
            self.logger.error(f"Error connecting to local target: {e}")
            return None

    def _compare_schemas(
        self,
        source: SchemaInfo,
        target: SchemaInfo,
        result: ValidationResult
    ):
        """Compare source and target schemas and populate findings."""

        # Compare tables
        missing_tables = source.tables - target.tables
        extra_tables = target.tables - source.tables

        for table in missing_tables:
            result.findings.append(ValidationFinding(
                severity=ValidationSeverity.ERROR,
                category="Table",
                message=f"Missing table: {table}",
                details="This table exists in source but not in target"
            ))

        for table in extra_tables:
            result.findings.append(ValidationFinding(
                severity=ValidationSeverity.INFO,
                category="Table",
                message=f"Extra table: {table}",
                details="This table exists in target but not in source"
            ))

        # Compare columns for common tables
        common_tables = source.tables & target.tables
        for table in common_tables:
            source_cols = source.columns.get(table, set())
            target_cols = target.columns.get(table, set())

            missing_cols = source_cols - target_cols
            extra_cols = target_cols - source_cols

            for col in missing_cols:
                result.findings.append(ValidationFinding(
                    severity=ValidationSeverity.WARNING,
                    category="Column",
                    message=f"Missing column: {table}[{col}]",
                    details="This column exists in source but not in target"
                ))

            # Extra columns are usually not a problem
            if len(extra_cols) > 5:
                result.findings.append(ValidationFinding(
                    severity=ValidationSeverity.INFO,
                    category="Column",
                    message=f"Target has {len(extra_cols)} extra columns in {table}",
                    details=f"Extra columns: {', '.join(list(extra_cols)[:5])}..."
                ))

        # Compare measures
        missing_measures = source.measures - target.measures
        extra_measures = target.measures - source.measures

        for measure in missing_measures:
            result.findings.append(ValidationFinding(
                severity=ValidationSeverity.WARNING,
                category="Measure",
                message=f"Missing measure: {measure}",
                details="This measure exists in source but not in target"
            ))

        if len(extra_measures) > 0 and len(extra_measures) <= 5:
            for measure in extra_measures:
                result.findings.append(ValidationFinding(
                    severity=ValidationSeverity.INFO,
                    category="Measure",
                    message=f"Extra measure: {measure}",
                    details="This measure exists in target but not in source"
                ))
        elif len(extra_measures) > 5:
            result.findings.append(ValidationFinding(
                severity=ValidationSeverity.INFO,
                category="Measure",
                message=f"Target has {len(extra_measures)} extra measures",
                details="Target model has additional measures not in source"
            ))

        # Compare relationship counts
        if source.relationships != target.relationships:
            diff = target.relationships - source.relationships
            result.findings.append(ValidationFinding(
                severity=ValidationSeverity.INFO,
                category="Relationship",
                message=f"Relationship count differs ({diff:+d})",
                details=f"Source: {source.relationships}, Target: {target.relationships}"
            ))

        # If no findings, add success message
        if not result.findings:
            result.findings.append(ValidationFinding(
                severity=ValidationSeverity.INFO,
                category="Validation",
                message="Schemas appear compatible",
                details="No significant differences detected"
            ))
