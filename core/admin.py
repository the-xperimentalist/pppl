from django.contrib import admin
from .models import (
    Project, Quote, CustomerGroup, MaterialGroup,
    AssemblyType, PackagingType, RawMaterial, MouldingMachineDetail,
    Assembly, AssemblyRawMaterial, ManufacturingPrintingCost, Packaging, Transport,
    QuoteTimeline, MaterialType, MouldingMachineType
)


@admin.register(Project)
class ProjectAdmin(admin.ModelAdmin):
    list_display = ['name', 'created_by', 'created_at', 'is_active', 'get_quotes_count']
    list_filter = ['is_active', 'created_at', 'created_by']
    search_fields = ['name', 'description']
    readonly_fields = ['created_at', 'updated_at']

    fieldsets = (
        ('Basic Information', {
            'fields': ('name', 'description', 'created_by', 'is_active')
        }),
        ('Timestamps', {
            'fields': ('created_at', 'updated_at'),
            'classes': ('collapse',)
        }),
    )

    def get_quotes_count(self, obj):
        return obj.get_quotes_count()
    get_quotes_count.short_description = 'Quotes Count'


@admin.register(CustomerGroup)
class CustomerGroupAdmin(admin.ModelAdmin):
    list_display = ['name', 'value', 'is_active', 'created_by', 'created_at']
    list_filter = ['is_active', 'created_at']
    search_fields = ['name', 'value', 'description']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(MaterialGroup)
class MaterialGroupAdmin(admin.ModelAdmin):
    list_display = ['name', 'value', 'is_active', 'created_by', 'created_at']
    list_filter = ['is_active', 'created_at']
    search_fields = ['name', 'value', 'description']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(AssemblyType)
class AssemblyTypeAdmin(admin.ModelAdmin):
    list_display = ['name', 'value', 'is_active', 'created_by', 'created_at']
    list_filter = ['is_active', 'created_at']
    search_fields = ['name', 'value', 'description']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(PackagingType)
class PackagingTypeAdmin(admin.ModelAdmin):
    list_display = ['name', 'value', 'is_active', 'created_by', 'created_at']
    list_filter = ['is_active', 'created_at']
    search_fields = ['name', 'value', 'description']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(Quote)
class QuoteAdmin(admin.ModelAdmin):
    list_display = ['name', 'project', 'client_name', 'status', 'get_completion_percentage', 'created_by', 'created_at']
    list_filter = ['status', 'created_at', 'project', 'created_by']
    search_fields = ['name', 'client_name', 'part_number', 'part_name', 'sap_number']
    readonly_fields = ['created_at', 'updated_at', 'get_completion_percentage']

    fieldsets = (
        ('Quote Definition', {
            'fields': ('project', 'name', 'client_group', 'client_name', 'status', 'notes')
        }),
        ('Part Details', {
            'fields': ('sap_number', 'part_number', 'part_name', 'amendment_number', 'description', 'quantity')
        }),
        ('Section Completion', {
            'fields': (
                'quote_definition_complete', 'raw_material_complete', 'moulding_machine_complete',
                'assembly_complete', 'packaging_complete', 'transport_complete'
            )
        }),
        ('Metadata', {
            'fields': ('created_by', 'created_at', 'updated_at'),
            'classes': ('collapse',)
        }),
    )

    def get_completion_percentage(self, obj):
        return f"{obj.get_completion_percentage()}%"
    get_completion_percentage.short_description = 'Completion'


@admin.register(RawMaterial)
class RawMaterialAdmin(admin.ModelAdmin):
    list_display = ['material_name', 'grade', 'material_type', 'rm_code', 'quote', 'rm_rate', 'created_at']
    list_filter = ['material_type', 'created_at']
    search_fields = ['material_name', 'grade', 'rm_code', 'quote__name']
    readonly_fields = ['created_at', 'updated_at', 'net_weight', 'gross_weight', 'rm_cost']

    fieldsets = (
        ('Quote', {
            'fields': ('quote',)
        }),
        ('Material Information', {
            'fields': ('material_type', 'material_name', 'grade', 'rm_code', 'unit_of_measurement')
        }),
        ('Pricing', {
            'fields': ('rm_rate', 'frozen_rate')
        }),
        ('Weight', {
            'fields': ('part_weight', 'runner_weight', 'net_weight', 'gross_weight')
        }),
        ('Additional Costs', {
            'fields': ('process_losses', 'purging_loss_cost', 'icc_percentage')
        }),
        ('Calculated', {
            'fields': ('rm_cost',)
        }),
        ('Timestamps', {
            'fields': ('created_at', 'updated_at')
        }),
    )

@admin.register(MouldingMachineDetail)
class MouldingMachineDetailAdmin(admin.ModelAdmin):
    list_display = ['quote', 'moulding_machine_type', 'cavity', 'machine_tonnage', 'cycle_time',
                    'efficiency', 'shift_rate', 'mtc_count', 'created_at']
    list_filter = ['moulding_machine_type', 'created_at']
    search_fields = ['quote__name', 'moulding_machine_type__name']
    readonly_fields = ['created_at', 'updated_at', 'number_of_parts_per_shift', 'mtc_cost', 'conversion_cost']

    fieldsets = (
        ('Quote', {
            'fields': ('quote',)
        }),
        ('Machine Type', {
            'fields': ('moulding_machine_type',)
        }),
        ('Machine Details', {
            'fields': ('cavity', 'machine_tonnage', 'cycle_time', 'efficiency')
        }),
        ('Shift and Cost Details', {
            'fields': ('shift_rate', 'shift_rate_for_mtc', 'mtc_count')
        }),
        ('Cost Percentages', {
            'fields': ('rejection_percentage', 'overhead_percentage', 'maintenance_percentage', 'profit_percentage')
        }),
        ('Calculated', {
            'fields': ('number_of_parts_per_shift', 'mtc_cost', 'conversion_cost')
        }),
        ('Timestamps', {
            'fields': ('created_at', 'updated_at')
        }),
    )


@admin.register(Assembly)
class AssemblyAdmin(admin.ModelAdmin):
    list_display = ['quote', 'assembly_type', 'total_assembly_rm_cost',
                    'total_manufacturing_printing_cost', 'total_assembly_cost']
    list_filter = ['assembly_type', 'quote']
    readonly_fields = ['total_assembly_rm_cost', 'total_manufacturing_printing_cost',
                      'profit_cost', 'rejection_cost', 'inspection_handling_cost',
                      'total_assembly_cost', 'created_at', 'updated_at']


@admin.register(AssemblyRawMaterial)
class AssemblyRawMaterialAdmin(admin.ModelAdmin):
    list_display = ['description', 'assembly', 'production_quantity', 'production_weight',
                    'cost_per_unit', 'total_cost']
    list_filter = ['assembly']
    readonly_fields = ['total_cost', 'created_at', 'updated_at']


@admin.register(ManufacturingPrintingCost)
class ManufacturingPrintingCostAdmin(admin.ModelAdmin):
    list_display = ['process', 'assembly', 'mc_tonnage', 'mc_rate_per_hour', 'cycle_time', 'per_cost']
    list_filter = ['assembly']
    readonly_fields = ['per_cost', 'created_at', 'updated_at']


@admin.register(Packaging)
class PackagingAdmin(admin.ModelAdmin):
    list_display = ['quote', 'packaging_type', 'lifecycle', 'cost', 'maintenance_percentage',
                    'part_per_polybag', 'total_packaging_cost']
    list_filter = ['packaging_type', 'quote']
    readonly_fields = ['length', 'breadth', 'height', 'polybag_length', 'polybag_width',
                      'maintenance_cost', 'total_cost', 'cost_per_part', 'polybag_cost',
                      'total_packaging_cost', 'created_at', 'updated_at']

    fieldsets = (
        ('Packaging Type', {
            'fields': ('quote', 'packaging_type')
        }),
        ('Fixed Dimensions (Auto-set)', {
            'fields': ('length', 'breadth', 'height', 'polybag_length', 'polybag_width'),
            'classes': ('collapse',)
        }),
        ('User Inputs', {
            'fields': ('lifecycle', 'cost', 'maintenance_percentage', 'part_per_polybag')
        }),
        ('Calculated Values', {
            'fields': ('maintenance_cost', 'total_cost', 'cost_per_part', 'polybag_cost', 'total_packaging_cost'),
            'classes': ('collapse',)
        }),
    )


@admin.register(Transport)
class TransportAdmin(admin.ModelAdmin):
    list_display = ['quote', 'packaging', 'transport_length', 'transport_breadth', 'transport_height',
                    'total_boxes', 'total_parts_per_trip', 'trip_cost_per_part']
    list_filter = ['quote', 'packaging']
    readonly_fields = ['boxes_on_length', 'boxes_on_breadth', 'boxes_on_height', 'total_boxes',
                      'total_parts_per_trip', 'trip_cost_per_part', 'created_at', 'updated_at']

    fieldsets = (
        ('Basic Info', {
            'fields': ('quote', 'packaging')
        }),
        ('Transport Dimensions', {
            'fields': ('transport_length', 'transport_breadth', 'transport_height')
        }),
        ('Cost & Parts', {
            'fields': ('trip_cost', 'parts_per_box')
        }),
        ('Calculated Values', {
            'fields': ('boxes_on_length', 'boxes_on_breadth', 'boxes_on_height', 'total_boxes',
                      'total_parts_per_trip', 'trip_cost_per_part'),
            'classes': ('collapse',)
        }),
    )


@admin.register(QuoteTimeline)
class QuoteTimelineAdmin(admin.ModelAdmin):
    list_display = ['quote', 'activity_type', 'user', 'created_at']
    list_filter = ['activity_type', 'created_at', 'quote']
    search_fields = ['quote__name', 'description']
    readonly_fields = ['created_at']

    fieldsets = (
        ('Timeline Entry', {
            'fields': ('quote', 'activity_type', 'description', 'user', 'attachment')
        }),
        ('Timestamp', {
            'fields': ('created_at',),
            'classes': ('collapse',)
        }),
    )


@admin.register(MaterialType)
class MaterialTypeAdmin(admin.ModelAdmin):
    list_display = ['raw_material_name', 'raw_material_grade', 'raw_material_rate',
                    'customer_group', 'is_active', 'created_by']
    list_filter = ['customer_group', 'is_active', 'created_at']
    search_fields = ['raw_material_name', 'raw_material_grade']
    readonly_fields = ['created_at', 'updated_at']

    fieldsets = (
        ('Customer Group', {
            'fields': ('customer_group',)
        }),
        ('Material Details', {
            'fields': ('raw_material_name', 'raw_material_grade', 'raw_material_rate')
        }),
        ('Status', {
            'fields': ('is_active', 'created_by')
        }),
    )


@admin.register(MouldingMachineType)
class MouldingMachineTypeAdmin(admin.ModelAdmin):
    list_display = ['name', 'customer_group', 'shift_rate', 'shift_rate_for_mtc', 'mtc_count',
                    'is_active', 'created_by']
    list_filter = ['customer_group', 'is_active', 'created_at']
    search_fields = ['name']
    readonly_fields = ['created_at', 'updated_at', 'mtc_cost']

    fieldsets = (
        ('Customer Group', {
            'fields': ('customer_group',)
        }),
        ('Machine Details', {
            'fields': ('name', 'shift_rate', 'shift_rate_for_mtc', 'mtc_count')
        }),
        ('Calculated', {
            'fields': ('mtc_cost',)
        }),
        ('Status', {
            'fields': ('is_active', 'created_by')
        }),
    )
