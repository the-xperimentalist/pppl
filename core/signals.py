from django.db.models.signals import post_save
from django.dispatch import receiver

from decimal import Decimal
from .models import (
    MaterialType, MouldingMachineType, AssemblyType,
    RawMaterial, MouldingMachineDetail, Assembly,
    CustomerGroup, PackagingType
)


@receiver(post_save, sender=CustomerGroup)
def create_default_packaging_types(sender, instance, created, **kwargs):
    """Create default packaging types when a customer group is created"""
    if created:
        default_packaging_types = [
            {
                'name': 'PP Box',
                'packaging_category': 'pp_box',
                'default_length': 600,
                'default_breadth': 400,
                'default_height': 250,
                'default_polybag_length': 16,
                'default_polybag_width': 20,
            },
            {
                'name': 'PP Box by Partition',
                'packaging_category': 'pp_box_partition',
                'default_length': 600,
                'default_breadth': 400,
                'default_height': 250,
                'default_polybag_length': 16,
                'default_polybag_width': 20,
            },
            {
                'name': 'CG Box',
                'packaging_category': 'cg_box',
                'default_length': 600,
                'default_breadth': 400,
                'default_height': 250,
                'default_polybag_length': 16,
                'default_polybag_width': 20,
            },
            {
                'name': 'Bin',
                'packaging_category': 'bin',
                'default_length': 600,
                'default_breadth': 400,
                'default_height': 250,
                'default_polybag_length': 16,
                'default_polybag_width': 20,
            },
        ]

        for pkg_type in default_packaging_types:
            PackagingType.objects.create(
                customer_group=instance,
                **pkg_type
            )

# =============================================================================
# AUTO-UPDATE QUOTE VALUES WHEN CONFIG TYPES CHANGE
# Create this file: core/signals.py
# =============================================================================


@receiver(post_save, sender=MaterialType)
def update_quotes_on_material_type_change(sender, instance, **kwargs):
    """
    When a MaterialType is updated, update all RawMaterials that reference it
    """
    # Find all RawMaterials that use this MaterialType
    raw_materials = RawMaterial.objects.filter(material_type=instance)

    for rm in raw_materials:
        # Update the fields that come from MaterialType
        rm.material_name = instance.raw_material_name
        rm.grade = instance.raw_material_grade
        rm.rm_code = instance.raw_material_code
        rm.rm_rate = instance.raw_material_rate
        rm.save()

        # Log the update to quote timeline
        from .models import QuoteTimeline
        QuoteTimeline.add_entry(
            quote=rm.quote,
            user=instance.created_by or rm.quote.created_by,
            description=f"Raw material '{rm.material_name}' auto-updated from Material Type template changes",
            activity_type="raw_material_auto_updated"
        )


@receiver(post_save, sender=MouldingMachineType)
def update_quotes_on_machine_type_change(sender, instance, **kwargs):
    """
    When a MouldingMachineType is updated, update all MouldingMachineDetails that reference it
    """
    # Find all MouldingMachineDetails that use this MouldingMachineType
    machines = MouldingMachineDetail.objects.filter(moulding_machine_type=instance)

    for machine in machines:
        # Update the fields that come from MouldingMachineType
        machine.shift_rate = instance.shift_rate
        machine.shift_rate_for_mtc = instance.shift_rate_for_mtc
        machine.mtc_count = instance.mtc_count
        machine.save()

        # Log the update to quote timeline
        from .models import QuoteTimeline
        QuoteTimeline.add_entry(
            quote=machine.quote,
            user=instance.created_by or machine.quote.created_by,
            description=f"Moulding machine (Cavity: {machine.cavity}) auto-updated from Machine Type template changes",
            activity_type="machine_auto_updated"
        )


@receiver(post_save, sender=AssemblyType)
def update_quotes_on_assembly_type_change(sender, instance, **kwargs):
    """
    When an AssemblyType is updated, update all Assemblies that reference it

    Note: AssemblyType mainly stores metadata (name, value, description).
    Most Assembly costs are entered manually, so this signal primarily updates
    the assembly name if the user wants to keep it in sync.
    """
    # Find all Assemblies that use this AssemblyType
    assemblies = Assembly.objects.filter(assembly_type_config=instance)

    for assembly in assemblies:
        # Update the name to match the AssemblyType name
        # Only update if the assembly name matches the old type name
        # This prevents overwriting custom assembly names
        if assembly.name == instance.name or not assembly.name:
            assembly.name = instance.name
            assembly.save()

            # Log the update to quote timeline
            from .models import QuoteTimeline
            QuoteTimeline.add_entry(
                quote=assembly.quote,
                user=instance.created_by or assembly.quote.created_by,
                description=f"Assembly '{assembly.name}' auto-updated from Assembly Type template changes",
                activity_type="assembly_auto_updated"
            )
