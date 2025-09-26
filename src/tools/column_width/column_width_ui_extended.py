    
    def _on_tree_double_click(self, event):
        """Handle double-click for per-visual configuration"""
        item = self.visual_tree.identify('item', event.x, event.y)
        if item:
            tags = self.visual_tree.item(item, 'tags')
            if tags:
                visual_id = tags[0]
                visual_info = next((v for v in self.visuals_info if v.visual_id == visual_id), None)
                if visual_info:
                    self._show_per_visual_config_dialog(visual_info)
    
    def _show_per_visual_config_dialog(self, visual_info: VisualInfo):
        """Show per-visual configuration dialog"""
        config_window = tk.Toplevel(self.main_app.root)
        config_window.title(f"Configure: {visual_info.visual_name}")
        config_window.geometry("500x600")
        config_window.resizable(True, True)
        config_window.transient(self.main_app.root)
        config_window.grab_set()
        
        # Center window
        config_window.geometry(f"+{self.main_app.root.winfo_rootx() + 100}+{self.main_app.root.winfo_rooty() + 100}")
        
        # Main frame
        main_frame = ttk.Frame(config_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(header_frame, text=f"⚙️ Configure: {visual_info.visual_name}", 
                 font=('Segoe UI', 14, 'bold')).pack(anchor=tk.W)
        
        visual_type = "Table" if visual_info.visual_type == VisualType.TABLE else "Matrix"
        ttk.Label(header_frame, text=f"Type: {visual_type} | Page: {visual_info.page_name}", 
                 font=('Segoe UI', 9),
                 foreground=AppConstants.COLORS['text_secondary']).pack(anchor=tk.W, pady=(5, 0))
        
        # Get or create per-visual config variables
        if visual_info.visual_id not in self.visual_config_vars:
            self._create_per_visual_config_vars(visual_info.visual_id)
        
        visual_vars = self.visual_config_vars[visual_info.visual_id]
        
        # Configuration sections
        self._setup_per_visual_width_controls(main_frame, visual_vars)
        
        # Field preview
        self._setup_field_preview(main_frame, visual_info)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Button(button_frame, text="⚙️ Use for This Visual",
                  command=lambda: self._apply_per_visual_config(visual_info.visual_id, config_window),
                  style='Action.TButton').pack(side=tk.LEFT)
        
        ttk.Button(button_frame, text="🌐 Copy to Global",
                  command=lambda: self._copy_to_global_config(visual_info.visual_id),
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(button_frame, text="❌ Cancel",
                  command=config_window.destroy,
                  style='Secondary.TButton').pack(side=tk.RIGHT)
    
    def _create_per_visual_config_vars(self, visual_id: str):
        """Create per-visual configuration variables"""
        self.visual_config_vars[visual_id] = {
            'categorical_preset': tk.StringVar(value=self.global_categorical_preset_var.get()),
            'categorical_custom': tk.IntVar(value=self.global_categorical_custom_var.get()),
            'measure_preset': tk.StringVar(value=self.global_measure_preset_var.get()),
            'measure_custom': tk.IntVar(value=self.global_measure_custom_var.get()),
            'max_width': tk.IntVar(value=self.global_max_width_var.get()),
            'min_width': tk.IntVar(value=self.global_min_width_var.get())
        }
    
    def _setup_per_visual_width_controls(self, parent, visual_vars):
        """Setup width controls for per-visual configuration"""
        config_frame = ttk.LabelFrame(parent, text="Width Configuration", padding="15")
        config_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Categorical and measure settings
        columns_frame = ttk.Frame(config_frame)
        columns_frame.pack(fill=tk.X)
        
        # Categorical columns
        cat_frame = ttk.LabelFrame(columns_frame, text="📊 Categorical", padding="10")
        cat_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self._setup_width_controls_compact(cat_frame, "categorical", 
                                          visual_vars['categorical_preset'], 
                                          visual_vars['categorical_custom'])
        
        # Measure columns
        measure_frame = ttk.LabelFrame(columns_frame, text="📈 Measures", padding="10")
        measure_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(5, 0))
        
        self._setup_width_controls_compact(measure_frame, "measure", 
                                          visual_vars['measure_preset'], 
                                          visual_vars['measure_custom'])
        
        # Limits
        limits_frame = ttk.Frame(config_frame)
        limits_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(limits_frame, text="🎯 Limits:", 
                 font=('Segoe UI', 9, 'bold')).pack(side=tk.LEFT)
        
        ttk.Label(limits_frame, text="Min:").pack(side=tk.LEFT, padx=(10, 2))
        min_spin = ttk.Spinbox(limits_frame, from_=30, to=200, width=6, textvariable=visual_vars['min_width'])
        min_spin.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(limits_frame, text="Max:").pack(side=tk.LEFT, padx=(0, 2))
        max_spin = ttk.Spinbox(limits_frame, from_=100, to=500, width=6, textvariable=visual_vars['max_width'])
        max_spin.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(limits_frame, text="px").pack(side=tk.LEFT)
    
    def _setup_field_preview(self, parent, visual_info: VisualInfo):
        """Setup field preview for per-visual configuration"""
        preview_frame = ttk.LabelFrame(parent, text="Field Preview", padding="15")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Field list
        field_tree = ttk.Treeview(preview_frame, height=8)
        field_tree.pack(fill=tk.BOTH, expand=True)
        
        field_tree['columns'] = ('type', 'current_width')
        field_tree.heading('#0', text='Field Name')
        field_tree.heading('type', text='Type')
        field_tree.heading('current_width', text='Current Width')
        
        field_tree.column('#0', width=200)
        field_tree.column('type', width=100)
        field_tree.column('current_width', width=100)
        
        # Populate fields
        for field in visual_info.fields:
            field_type = "📊 C" if field.field_type == FieldType.CATEGORICAL else "📈 M"
            current = f"{field.current_width:.0f}px" if field.current_width else "Not set"
            
            field_tree.insert('', 'end',
                            text=field.display_name,
                            values=(field_type, current))
    
    def _apply_per_visual_config(self, visual_id: str, config_window):
        """Apply per-visual configuration and switch to it"""
        self.current_selected_visual = visual_id
        self._switch_to_per_visual_config(visual_id)
        
        # Update config mode label
        visual_info = next((v for v in self.visuals_info if v.visual_id == visual_id), None)
        if visual_info:
            self.config_mode_label.config(text=f"Per-Visual: {visual_info.visual_name}")
        
        config_window.destroy()
        self.log_message(f"⚙️ Switched to per-visual configuration for {visual_info.visual_name if visual_info else visual_id}")
    
    def _copy_to_global_config(self, visual_id: str):
        """Copy per-visual configuration to global settings"""
        if visual_id in self.visual_config_vars:
            visual_vars = self.visual_config_vars[visual_id]
            
            # Copy values to global variables
            self.global_categorical_preset_var.set(visual_vars['categorical_preset'].get())
            self.global_categorical_custom_var.set(visual_vars['categorical_custom'].get())
            self.global_measure_preset_var.set(visual_vars['measure_preset'].get())
            self.global_measure_custom_var.set(visual_vars['measure_custom'].get())
            self.global_max_width_var.set(visual_vars['max_width'].get())
            self.global_min_width_var.set(visual_vars['min_width'].get())
            
            self.log_message("🌐 Per-visual configuration copied to global settings")
    
    def _switch_to_global_config(self):
        """Switch configuration UI to global settings"""
        self.categorical_preset_var = self.global_categorical_preset_var
        self.categorical_custom_var = self.global_categorical_custom_var
        self.measure_preset_var = self.global_measure_preset_var
        self.measure_custom_var = self.global_measure_custom_var
        self.max_width_var = self.global_max_width_var
        self.min_width_var = self.global_min_width_var
        
        self.config_mode_label.config(text="Global Settings (All Visuals)")
    
    def _switch_to_per_visual_config(self, visual_id: str):
        """Switch configuration UI to per-visual settings"""
        if visual_id in self.visual_config_vars:
            visual_vars = self.visual_config_vars[visual_id]
            
            self.categorical_preset_var = visual_vars['categorical_preset']
            self.categorical_custom_var = visual_vars['categorical_custom']
            self.measure_preset_var = visual_vars['measure_preset']
            self.measure_custom_var = visual_vars['measure_custom']
            self.max_width_var = visual_vars['max_width']
            self.min_width_var = visual_vars['min_width']
    
    def _apply_scale_intelligence(self):
        """Apply scale intelligence to measure fields"""
        if not self.engine:
            self.show_error("No Data", "Please scan visuals first.")
            return
        
        from tools.column_width.column_width_core import ScaleConfiguration, DataScale, DataMagnitude
        
        # Create scale configuration from UI
        scale_config = ScaleConfiguration(
            typical_scale=DataScale(self.scale_var.get()),
            use_abbreviations=self.use_abbreviations_var.get(),
            decimal_places=self.decimal_places_var.get(),
            currency_symbol=self.currency_symbol_var.get()
        )
        
        # Determine magnitude from scale
        scale_to_magnitude = {
            DataScale.ONES: DataMagnitude.SMALL,
            DataScale.TENS: DataMagnitude.SMALL,
            DataScale.HUNDREDS: DataMagnitude.SMALL,
            DataScale.THOUSANDS: DataMagnitude.MEDIUM,
            DataScale.MILLIONS: DataMagnitude.LARGE,
            DataScale.BILLIONS: DataMagnitude.XLARGE,
            DataScale.TRILLIONS: DataMagnitude.XLARGE
        }
        
        scale_config.magnitude = scale_to_magnitude.get(scale_config.typical_scale, DataMagnitude.MEDIUM)
        
        # Apply scale configuration to engine
        self.engine.apply_scale_configuration(scale_config)
        
        # Log the application
        scale_name = self.scale_var.get().title()
        format_desc = "abbreviated" if self.use_abbreviations_var.get() else "full"
        currency_desc = "with currency" if self.currency_symbol_var.get() else "without currency"
        
        self.log_message(f"🧠 Applied {scale_name} scale intelligence ({format_desc}, {currency_desc})")
        self.log_message(f"📈 Measure columns will be optimized for {scale_name}-scale data")
    
    def _get_current_config(self) -> WidthConfiguration:
        """Get current width configuration from UI - enhanced for per-visual support"""
        config = WidthConfiguration()
        
        # Use currently active configuration variables
        config.categorical_preset = WidthPreset(self.categorical_preset_var.get())
        config.categorical_custom = self.categorical_custom_var.get()
        config.measure_preset = WidthPreset(self.measure_preset_var.get())
        config.measure_custom = self.measure_custom_var.get()
        config.max_width = self.max_width_var.get()
        config.min_width = self.min_width_var.get()
        
        return config
    
    def _get_selected_visual_configs(self) -> Dict[str, WidthConfiguration]:
        """Get configuration for each selected visual"""
        configs = {}
        selected_ids = self._get_selected_visual_ids()
        
        for visual_id in selected_ids:
            if visual_id in self.visual_config_vars:
                # Use per-visual configuration
                visual_vars = self.visual_config_vars[visual_id]
                config = WidthConfiguration()
                config.categorical_preset = WidthPreset(visual_vars['categorical_preset'].get())
                config.categorical_custom = visual_vars['categorical_custom'].get()
                config.measure_preset = WidthPreset(visual_vars['measure_preset'].get())
                config.measure_custom = visual_vars['measure_custom'].get()
                config.max_width = visual_vars['max_width'].get()
                config.min_width = visual_vars['min_width'].get()
                configs[visual_id] = config
            else:
                # Use global configuration
                configs[visual_id] = self._get_global_config()
        
        return configs
    
    def _get_global_config(self) -> WidthConfiguration:
        """Get global width configuration"""
        config = WidthConfiguration()
        config.categorical_preset = WidthPreset(self.global_categorical_preset_var.get())
        config.categorical_custom = self.global_categorical_custom_var.get()
        config.measure_preset = WidthPreset(self.global_measure_preset_var.get())
        config.measure_custom = self.global_measure_custom_var.get()
        config.max_width = self.global_max_width_var.get()
        config.min_width = self.global_min_width_var.get()
        return config
