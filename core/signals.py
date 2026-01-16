from django.db.models.signals import post_save
from django.dispatch import receiver
from core.models import CustomerGroup, PackagingType


@receiver(post_save, sender=CustomerGroup)
def create_default_packaging_types(sender, instance, created, **kwargs):
    """Create default packaging types when a customer group is created"""
    if created:
        default_packaging_types = [
            {
                'name': 'PP Box',
                'packaging_type': 'pp_box',
                'default_length': 600,
                'default_breadth': 400,
                'default_height': 250,
                'default_polybag_length': 16,
                'default_polybag_width': 20,
            },
            {
                'name': 'PP Box by Partition',
                'packaging_type': 'pp_box_partition',
                'default_length': 600,
                'default_breadth': 400,
                'default_height': 250,
                'default_polybag_length': 16,
                'default_polybag_width': 20,
            },
            {
                'name': 'CG Box',
                'packaging_type': 'cg_box',
                'default_length': 600,
                'default_breadth': 400,
                'default_height': 250,
                'default_polybag_length': 16,
                'default_polybag_width': 20,
            },
            {
                'name': 'Bin',
                'packaging_type': 'bin',
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
