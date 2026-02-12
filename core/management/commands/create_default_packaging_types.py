from django.core.management.base import BaseCommand
from core.models import CustomerGroup, PackagingType


class Command(BaseCommand):
    help = 'Create default packaging types for existing customer groups'

    def handle(self, *args, **options):
        default_packaging_categorys = [
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
        
        customer_groups = CustomerGroup.objects.all()
        
        for customer_group in customer_groups:
            for pkg_type in default_packaging_categorys:
                # Check if packaging type already exists
                existing = PackagingType.objects.filter(
                    customer_group=customer_group,
                    name=pkg_type['name']
                ).exists()
                
                if not existing:
                    PackagingType.objects.create(
                        customer_group=customer_group,
                        **pkg_type
                    )
                    self.stdout.write(
                        self.style.SUCCESS(
                            f'Created "{pkg_type["name"]}" for {customer_group.name}'
                        )
                    )
                else:
                    self.stdout.write(
                        self.style.WARNING(
                            f'"{pkg_type["name"]}" already exists for {customer_group.name}'
                        )
                    )
        
        self.stdout.write(self.style.SUCCESS('Successfully created default packaging types'))