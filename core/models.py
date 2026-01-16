from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
from decimal import Decimal


class CustomerGroup(models.Model):
    """Customer Group configuration"""
    name = models.CharField(max_length=100, unique=True, default="")
    value = models.CharField(max_length=100, unique=True, default="", help_text="Unique identifier/code")
    description = models.TextField(blank=True, null=True, default="")
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='customer_groups')
    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ['name']
        verbose_name = 'Customer Group'
        verbose_name_plural = 'Customer Groups'

    def __str__(self):
        return f"{self.name} ({self.value})"


class MouldingMachineType(models.Model):
    """Moulding Machine Type - belongs to a customer group"""
    customer_group = models.ForeignKey(CustomerGroup, on_delete=models.CASCADE, related_name='moulding_machine_types',
                                       null=True, blank=True, default=None)

    # Machine type details
    name = models.CharField(max_length=200, default="", help_text="Name/identifier for this machine type")
    shift_rate = models.DecimalField(max_digits=18, decimal_places=8, default=0, help_text="Shift rate")
    shift_rate_for_mtc = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                             verbose_name="Shift Rate for MTC")
    mtc_count = models.IntegerField(default=0, verbose_name="MTC Count", help_text="Number of MTC")

    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='moulding_machine_types')
    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ['customer_group', 'name']
        verbose_name = 'Moulding Machine Type'
        verbose_name_plural = 'Moulding Machine Types'

    def __str__(self):
        if self.name:
            return f"{self.name} - {self.customer_group.name if self.customer_group else 'No Group'}"
        return f"Moulding Machine (Shift: {self.shift_rate}) - {self.customer_group.name if self.customer_group else 'No Group'}"

    @property
    def mtc_cost(self):
        """Calculate MTC cost = mtc_count * shift_rate_for_mtc"""
        from decimal import Decimal
        return float(Decimal(str(self.mtc_count)) * Decimal(str(self.shift_rate_for_mtc)))


class MaterialType(models.Model):
    """Material Type - belongs to a customer group"""
    customer_group = models.ForeignKey(CustomerGroup, on_delete=models.CASCADE, related_name='material_types',
                                       null=True, blank=True, default=None)

    # Material details
    raw_material_name = models.CharField(max_length=200, default="")
    raw_material_grade = models.CharField(max_length=100, default="")
    raw_material_code = models.CharField(max_length=100, default="", verbose_name="RM Code")
    raw_material_rate = models.DecimalField(max_digits=18, decimal_places=8, default=0)

    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='material_types')
    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ['customer_group', 'raw_material_name']
        verbose_name = 'Material Type'
        verbose_name_plural = 'Material Types'

    def __str__(self):
        return f"{self.raw_material_name} - {self.raw_material_grade} ({self.customer_group.name if self.customer_group else 'No Group'})"


class Project(models.Model):
    """
    Project model - each project belongs to a user and contains multiple quotes
    """
    name = models.CharField(max_length=200, default="")
    description = models.TextField(blank=True, null=True, default="")
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='projects')
    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Project'
        verbose_name_plural = 'Projects'

    def __str__(self):
        return self.name

    def get_quotes_count(self):
        """Return the number of quotes in this project"""
        return self.quotes.count()


class MaterialGroup(models.Model):
    """Material Group configuration - belongs to customer group"""
    customer_group = models.ForeignKey(CustomerGroup, on_delete=models.CASCADE, related_name='material_groups',
                                       null=True, blank=True, default=None)
    name = models.CharField(max_length=100, default="")
    value = models.CharField(max_length=100, default="", help_text="Unique identifier/code")
    description = models.TextField(blank=True, null=True, default="")
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='material_groups')
    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ['customer_group', 'name']
        verbose_name = 'Material Group'
        verbose_name_plural = 'Material Groups'

    def __str__(self):
        return f"{self.name} ({self.value}) - {self.customer_group.name if self.customer_group else 'No Group'}"


class Quote(models.Model):
    """
    Quote model - each quote belongs to a project
    By default, all quotes are visible to all users
    """
    STATUS_CHOICES = [
        ('in_progress', 'In Progress'),
        ('completed', 'Completed'),
        ('discarded', 'Discarded'),
    ]

    project = models.ForeignKey(Project, on_delete=models.CASCADE, related_name='quotes')

    # Quote Definition fields
    name = models.CharField(max_length=200, default="", help_text="Quote name/title")
    client_group = models.ForeignKey(CustomerGroup, on_delete=models.PROTECT, related_name='quotes')  # MANDATORY
    client_name = models.CharField(max_length=200, default="")
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='in_progress')
    notes = models.TextField(blank=True, null=True, default="")

    # Financial parameters
    handling_charge = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                         help_text="Fixed handling charge")
    profit_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                           help_text="Profit percentage")

    # Version tracking
    major_version = models.IntegerField(default=1, help_text="Major version (e.g., 1 in 1.5)")
    minor_version = models.IntegerField(default=0, help_text="Minor version (e.g., 5 in 1.5)")

    sap_number = models.CharField(max_length=100, blank=True, null=True, default="", verbose_name="SAP Number")
    part_number = models.CharField(max_length=100, default="")
    part_name = models.CharField(max_length=200, default="")
    amendment_number = models.CharField(max_length=50, blank=True, null=True, default="")
    description = models.TextField(blank=True, null=True, default="")
    quantity = models.IntegerField(default=1)

    # Tracking
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='quotes')
    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)

    # Section completion tracking
    quote_definition_complete = models.BooleanField(default=False)
    raw_material_complete = models.BooleanField(default=False)
    moulding_machine_complete = models.BooleanField(default=False)
    assembly_complete = models.BooleanField(default=False)
    packaging_complete = models.BooleanField(default=False)
    transport_complete = models.BooleanField(default=False)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Quote'
        verbose_name_plural = 'Quotes'

    def __str__(self):
        return f"{self.name} - {self.project.name} (v{self.get_version()})"

    def get_version(self):
        """Return version as string (e.g., '1.5')"""
        return f"{self.major_version}.{self.minor_version}"

    def increment_version(self, user, description="Quote updated", activity_type="quote_updated", create_timeline=True):
        """Increment minor version and log to timeline"""
        self.minor_version += 1
        self.save()

        # Add timeline entry
        if create_timeline:
            QuoteTimeline.add_entry(
                quote=self,
                activity_type=activity_type,
                description=f'{description} - Version updated to {self.get_version()}',
                user=user
            )

    def mark_completed(self, user):
        """Mark quote as completed and bump to next major version"""
        if self.status == 'in_progress':
            self.status = 'completed'
            self.major_version += 1
            self.minor_version = 0
            self.save()

            # Add timeline entry
            QuoteTimeline.add_entry(
                quote=self,
                activity_type='section_completed',
                description=f'Quote marked as completed - Version updated to {self.get_version()}',
                user=user
            )
            return True
        return False

    def reopen_quote(self, user):
        """Reopen completed quote for editing"""
        if self.status == 'completed':
            self.status = 'in_progress'
            self.major_version += 1
            self.minor_version = 0
            self.save()

            # Add timeline entry
            QuoteTimeline.add_entry(
                quote=self,
                activity_type='quote_updated',
                description=f'Quote reopened for editing - Version updated to {self.get_version()}',
                user=user
            )
            return True
        return False

    def discard_quote(self, user):
        """Discard a completed quote"""
        if self.status == 'completed':
            self.status = 'discarded'
            self.save()

            # Add timeline entry
            QuoteTimeline.add_entry(
                quote=self,
                activity_type='quote_updated',
                description=f'Quote discarded - Final version: {self.get_version()}',
                user=user
            )
            return True
        return False

    def can_edit_sections(self):
        """Check if quote sections can be edited"""
        return self.status == 'in_progress'

    def is_complete(self):
        """Check if all sections are complete"""
        return all([
            self.quote_definition_complete,
            self.raw_material_complete,
            self.moulding_machine_complete,
            self.assembly_complete,
            self.packaging_complete,
            self.transport_complete,
        ])

    def get_completion_percentage(self):
        """Calculate completion percentage"""
        sections = [
            self.quote_definition_complete,
            self.raw_material_complete,
            self.moulding_machine_complete,
            self.assembly_complete,
            self.packaging_complete,
            self.transport_complete,
        ]
        completed = sum(sections)
        return int((completed / len(sections)) * 100)

    def get_total_raw_material_cost(self):
        """Calculate total raw material cost"""
        return sum(rm.rm_cost or 0 for rm in self.raw_materials.all())

    def get_total_conversion_cost(self):
        """Calculate total conversion cost from all moulding machines"""
        return sum(mm.conversion_cost for mm in self.moulding_machines.all())

    def get_total_assembly_cost(self):
        """Calculate total assembly cost from all assemblies"""
        return sum(assembly.total_assembly_cost for assembly in self.assemblies.all())

    def get_total_packaging_cost(self):
        """Calculate total packaging cost from all packagings"""
        return sum(packaging.total_packaging_cost for packaging in self.packagings.all())

    def get_total_transport_cost(self):
        """Calculate total transport cost per part from all transports"""
        return sum(transport.trip_cost_per_part for transport in self.transports.all())

    def get_base_cost(self):
        """Calculate base cost before profit and handling charge"""
        return (
            Decimal(self.get_total_raw_material_cost()) +
            Decimal(self.get_total_conversion_cost()) +
            Decimal(self.get_total_assembly_cost()) +
            Decimal(self.get_total_packaging_cost()) +
            Decimal(self.get_total_transport_cost())
        )

    def get_profit_amount(self):
        """Calculate profit amount"""
        from decimal import Decimal
        base_cost = Decimal(str(self.get_base_cost()))
        profit_percentage = Decimal(str(self.profit_percentage))
        return base_cost * (profit_percentage / Decimal('100'))

    def get_grand_total(self):
        """Calculate grand total including profit and handling charge"""
        from decimal import Decimal
        base_cost = Decimal(str(self.get_base_cost()))
        profit = Decimal(str(self.get_profit_amount()))
        handling = Decimal(str(self.handling_charge))
        return base_cost + profit + handling


class RawMaterial(models.Model):
    """Raw Material for a quote"""
    UNIT_CHOICES = [
        ('kg', 'Kilogram (kg)'),
        ('gm', 'Gram (gm)'),
        ('ton', 'Ton'),
    ]

    quote = models.ForeignKey(Quote, on_delete=models.CASCADE, related_name='raw_materials')
    material_type = models.ForeignKey(MaterialType, on_delete=models.SET_NULL, null=True, blank=True,
                                      related_name='raw_materials', help_text="Select material type to auto-fill details")

    # Material details
    material_name = models.CharField(max_length=200, default="")
    grade = models.CharField(max_length=100, blank=True, null=True, default="")
    rm_code = models.CharField(max_length=100, default="", verbose_name="RM Code")
    unit_of_measurement = models.CharField(max_length=10, choices=UNIT_CHOICES, default='kg')

    # Pricing
    frozen_rate = models.DecimalField(max_digits=18, decimal_places=8, null=True, blank=True,
                                     help_text="Frozen rate if applicable (per kg)")
    rm_rate = models.DecimalField(max_digits=18, decimal_places=8, default=0, verbose_name="RM Rate (per kg)")

    # Weight details
    part_weight = models.DecimalField(max_digits=18, decimal_places=8, default=0)
    runner_weight = models.DecimalField(max_digits=18, decimal_places=8, default=0)

    # Additional costs
    process_losses = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                        help_text="Process losses cost")
    purging_loss_cost = models.DecimalField(max_digits=18, decimal_places=8, default=0)
    icc_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                        verbose_name="ICC %", help_text="ICC percentage")

    # Cost percentages
    rejection_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                              help_text="Rejection percentage")
    overhead_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                             help_text="Overhead percentage")
    maintenance_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                                help_text="Maintenance percentage")
    profit_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                           help_text="Profit percentage")

    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['created_at']
        verbose_name = 'Raw Material'
        verbose_name_plural = 'Raw Materials'

    def __str__(self):
        return f"{self.material_name} - {self.quote.name}"

    @property
    def gross_weight(self):
        """Calculate gross weight (part + runner) in the selected unit"""
        from decimal import Decimal
        return float(Decimal(str(self.part_weight)) + Decimal(str(self.runner_weight)))

    @property
    def gross_weight_in_grams(self):
        """Convert gross weight to grams based on unit of measurement"""
        from decimal import Decimal

        gross = Decimal(str(self.gross_weight))

        # Convert to grams based on unit
        if self.unit_of_measurement == 'kg':
            # 1 kg = 1000 grams
            return float(gross * Decimal('1000'))
        elif self.unit_of_measurement == 'ton':
            # 1 ton = 1,000,000 grams
            return float(gross * Decimal('1000000'))
        else:  # 'gm'
            # Already in grams
            return float(gross)

    @property
    def effective_rate_per_kg(self):
        """Get effective rate per kg (frozen rate if available, otherwise rm_rate)"""
        if self.frozen_rate is not None and self.frozen_rate > 0:
            return float(self.frozen_rate)
        return float(self.rm_rate)

    @property
    def effective_rate_per_gram(self):
        """Convert effective rate from per kg to per gram (divide by 1000)"""
        from decimal import Decimal
        rate_per_kg = Decimal(str(self.effective_rate_per_kg))
        # 1 kg = 1000 grams, so rate per gram = rate per kg / 1000
        return float(rate_per_kg / Decimal('1000'))

    @property
    def base_rm_cost(self):
        """Calculate base RM cost using gross weight in grams and rate per gram"""
        from decimal import Decimal

        # Use gross weight in grams
        weight_in_grams = Decimal(str(self.gross_weight_in_grams))

        # Rate per gram (converted from per kg)
        rate_per_gram = Decimal(str(self.effective_rate_per_gram))

        # Base cost = weight in grams × rate per gram + additional costs
        base = (weight_in_grams * rate_per_gram) + Decimal(str(self.process_losses)) + Decimal(str(self.purging_loss_cost))

        # Add ICC percentage
        icc_cost = base * (Decimal(str(self.icc_percentage)) / Decimal('100'))

        return float(base + icc_cost)

    @property
    def rejection_cost(self):
        """Calculate rejection cost"""
        from decimal import Decimal
        return float(Decimal(str(self.base_rm_cost)) * (Decimal(str(self.rejection_percentage)) / Decimal('100')))

    @property
    def overhead_cost(self):
        """Calculate overhead cost"""
        from decimal import Decimal
        return float(Decimal(str(self.base_rm_cost)) * (Decimal(str(self.overhead_percentage)) / Decimal('100')))

    @property
    def maintenance_cost(self):
        """Calculate maintenance cost"""
        from decimal import Decimal
        return float(Decimal(str(self.base_rm_cost)) * (Decimal(str(self.maintenance_percentage)) / Decimal('100')))

    @property
    def profit_cost(self):
        """Calculate profit cost"""
        from decimal import Decimal
        return float(Decimal(str(self.base_rm_cost)) * (Decimal(str(self.profit_percentage)) / Decimal('100')))

    @property
    def rm_cost(self):
        """Calculate total RM cost"""
        from decimal import Decimal
        return float(
            Decimal(str(self.base_rm_cost)) +
            Decimal(str(self.rejection_cost)) +
            Decimal(str(self.overhead_cost)) +
            Decimal(str(self.maintenance_cost)) +
            Decimal(str(self.profit_cost))
        )


class AssemblyType(models.Model):
    """Assembly Type configuration - belongs to customer group"""
    customer_group = models.ForeignKey(CustomerGroup, on_delete=models.CASCADE, related_name='assembly_types',
                                       null=True, blank=True, default=None)
    name = models.CharField(max_length=100, default="")
    value = models.CharField(max_length=100, default="", help_text="Unique identifier/code")
    description = models.TextField(blank=True, null=True, default="")
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='assembly_types')
    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ['customer_group', 'name']
        verbose_name = 'Assembly Type'
        verbose_name_plural = 'Assembly Types'

    def __str__(self):
        return f"{self.name} ({self.value}) - {self.customer_group.name if self.customer_group else 'No Group'}"


class PackagingType(models.Model):
    """Packaging type configuration for customer groups"""
    PACKAGING_TYPE_CHOICES = [
        ('pp_box', 'PP Box'),
        ('pp_box_partition', 'PP Box by Partition'),
        ('cg_box', 'CG Box'),
        ('bin', 'Bin'),
        ('custom', 'Custom'),
    ]

    customer_group = models.ForeignKey(CustomerGroup, on_delete=models.CASCADE,
                                      related_name='packaging_types', null=True)
    name = models.CharField(max_length=100, help_text="Packaging type name")
    packaging_type = models.CharField(max_length=50, choices=PACKAGING_TYPE_CHOICES,
                                     default='pp_box')

    # Default dimensions
    default_length = models.DecimalField(max_digits=18, decimal_places=8, default=600,
                                        help_text="Default packaging length (mm)")
    default_breadth = models.DecimalField(max_digits=18, decimal_places=8, default=400,
                                         help_text="Default packaging breadth (mm)")
    default_height = models.DecimalField(max_digits=18, decimal_places=8, default=250,
                                        help_text="Default packaging height (mm)")
    default_polybag_length = models.DecimalField(max_digits=18, decimal_places=8, default=16,
                                                 help_text="Default polybag length")
    default_polybag_width = models.DecimalField(max_digits=18, decimal_places=8, default=20,
                                               help_text="Default polybag width")

    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['packaging_type', 'name']
        verbose_name = 'Packaging Type'
        verbose_name_plural = 'Packaging Types'
        unique_together = ['customer_group', 'name']

    def __str__(self):
        return f"{self.name} - {self.customer_group.name}"


class MouldingMachineDetail(models.Model):
    """Moulding Machine details for a quote"""
    quote = models.ForeignKey(Quote, on_delete=models.CASCADE, related_name='moulding_machines')
    moulding_machine_type = models.ForeignKey(MouldingMachineType, on_delete=models.SET_NULL, null=True, blank=True,
                                              related_name='moulding_machines',
                                              help_text="Select machine type to auto-fill details")

    # Machine details
    cavity = models.IntegerField(default=1)
    machine_tonnage = models.DecimalField(max_digits=18, decimal_places=8, default=0)
    cycle_time = models.DecimalField(max_digits=18, decimal_places=8, default=0, help_text="Cycle time in seconds")
    efficiency = models.DecimalField(max_digits=13, decimal_places=8, default=0, help_text="Efficiency percentage")

    # Shift and cost details
    shift_rate = models.DecimalField(max_digits=18, decimal_places=8, default=0)
    shift_rate_for_mtc = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                             verbose_name="Shift Rate for MTC")
    mtc_count = models.IntegerField(default=0, verbose_name="MTC Count", help_text="Number of MTC")

    # Cost percentages
    rejection_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                              help_text="Rejection percentage")
    overhead_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                             help_text="Overhead percentage")
    maintenance_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                                help_text="Maintenance percentage")
    profit_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                           help_text="Profit percentage")

    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['created_at']
        verbose_name = 'Moulding Machine Detail'
        verbose_name_plural = 'Moulding Machine Details'

    def __str__(self):
        return f"Machine {self.cavity} cavity - {self.quote.name}"

    @property
    def number_of_parts_per_shift(self):
        """Calculate number of parts per shift"""
        if self.cycle_time > 0 and self.efficiency > 0:
            # Assuming 8-hour shift = 28800 seconds
            shift_seconds = 28800
            effective_time = shift_seconds * (self.efficiency / 100)
            parts = (effective_time / self.cycle_time) * self.cavity
            return int(parts)
        return 0

    @property
    def number_of_mtc(self):
        """Calculate number of MTC (tool changes)"""
        return self.mtc_count

    @property
    def mtc_cost(self):
        """Calculate MTC cost = mtc_count * shift_rate_for_mtc"""
        from decimal import Decimal
        return float(Decimal(str(self.mtc_count)) * Decimal(str(self.shift_rate_for_mtc)))

    @property
    def base_conversion_cost(self):
        """Calculate base conversion cost per part before percentages"""
        parts_per_shift = self.number_of_parts_per_shift
        if parts_per_shift > 0:
            from decimal import Decimal
            # Base cost includes shift rate + MTC cost
            total_shift_cost = Decimal(str(self.shift_rate)) + Decimal(str(self.mtc_cost))
            return float(total_shift_cost / parts_per_shift)
        return 0

    @property
    def rejection_cost(self):
        """Calculate rejection cost"""
        from decimal import Decimal
        base = Decimal(str(self.base_conversion_cost))
        return float(base * (Decimal(str(self.rejection_percentage)) / Decimal('100')))

    @property
    def overhead_cost(self):
        """Calculate overhead cost"""
        from decimal import Decimal
        base = Decimal(str(self.base_conversion_cost))
        return float(base * (Decimal(str(self.overhead_percentage)) / Decimal('100')))

    @property
    def machine_maintenance_cost(self):
        """Calculate maintenance cost"""
        from decimal import Decimal
        base = Decimal(str(self.base_conversion_cost))
        return float(base * (Decimal(str(self.maintenance_percentage)) / Decimal('100')))

    @property
    def machine_profit_cost(self):
        """Calculate profit cost"""
        from decimal import Decimal
        base = Decimal(str(self.base_conversion_cost))
        return float(base * (Decimal(str(self.profit_percentage)) / Decimal('100')))

    @property
    def conversion_cost(self):
        """Calculate total conversion cost per part including all percentages"""
        from decimal import Decimal
        base = Decimal(str(self.base_conversion_cost))
        rejection = Decimal(str(self.rejection_cost))
        overhead = Decimal(str(self.overhead_cost))
        maintenance = Decimal(str(self.machine_maintenance_cost))
        profit = Decimal(str(self.machine_profit_cost))
        total = base + rejection + overhead + maintenance + profit
        return float(total)


class Assembly(models.Model):
    """Assembly for a quote"""
    quote = models.ForeignKey(Quote, on_delete=models.CASCADE, related_name='assemblies')
    assembly_type_config = models.ForeignKey(AssemblyType, on_delete=models.SET_NULL, null=True, blank=True,
                                             related_name='assemblies', help_text="Assembly type configuration")

    # Assembly details
    name = models.CharField(max_length=200, default="", help_text="Assembly name/identifier")
    remarks = models.TextField(blank=True, null=True, default="", help_text="Additional notes or remarks")
    manual_cost = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                     help_text="Manual assembly cost (used if automated)")
    other_cost = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                    help_text="Other miscellaneous costs")

    # Percentages and fixed costs
    profit_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0)
    rejection_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0)
    inspection_handling_cost = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                                   verbose_name="Inspection & Handling Cost",
                                                   help_text="Fixed cost for inspection and handling")

    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['created_at']
        verbose_name = 'Assembly'
        verbose_name_plural = 'Assemblies'

    def __str__(self):
        if self.name:
            return f"{self.name} - {self.quote.name}"
        return f"Assembly - {self.quote.name}"

    def calculate_costs(self):
        """Calculate all assembly costs"""
        from decimal import Decimal

        # Get total assembly RM cost
        total_assembly_rm_cost = sum(
            Decimal(str(rm.total_cost)) for rm in self.assembly_raw_materials.all()
        )

        # Get total manufacturing/printing cost
        total_manufacturing_printing_cost = sum(
            Decimal(str(cost.per_cost)) for cost in self.manufacturing_printing_costs.all()
        )

        # Base cost is sum of manual cost, assembly RM, and manufacturing/printing
        base_cost = (Decimal(str(self.manual_cost)) +
                    total_assembly_rm_cost +
                    total_manufacturing_printing_cost)

        # Calculate percentage-based costs
        profit_cost = base_cost * (Decimal(str(self.profit_percentage)) / Decimal('100'))
        rejection_cost = base_cost * (Decimal(str(self.rejection_percentage)) / Decimal('100'))

        # Inspection & handling is now a fixed cost (not percentage-based)
        inspection_handling_cost = Decimal(str(self.inspection_handling_cost))

        # Total assembly cost
        total = (base_cost +
                Decimal(str(self.other_cost)) +
                profit_cost +
                rejection_cost +
                inspection_handling_cost)

        return {
            'base_cost': float(base_cost),
            'profit_cost': float(profit_cost),
            'rejection_cost': float(rejection_cost),
            'inspection_handling_cost': float(inspection_handling_cost),
            'total_assembly_cost': float(total),
        }

    @property
    def base_cost(self):
        """Get base cost"""
        return self.calculate_costs()['base_cost']

    @property
    def profit_cost(self):
        """Get profit cost"""
        return self.calculate_costs()['profit_cost']

    @property
    def rejection_cost(self):
        """Get rejection cost"""
        return self.calculate_costs()['rejection_cost']

    @property
    def inspection_handling_cost_calculated(self):
        """Get inspection & handling cost (already a fixed value)"""
        return self.calculate_costs()['inspection_handling_cost']

    @property
    def total_assembly_cost(self):
        """Get total assembly cost"""
        return self.calculate_costs()['total_assembly_cost']

    def save(self, *args, **kwargs):
        """Override save to trigger cost recalculation"""
        super().save(*args, **kwargs)


class AssemblyRawMaterial(models.Model):
    """Raw Material for an assembly"""
    UNIT_CHOICES = [
        ('kg', 'Kilogram (kg)'),
        ('gm', 'Gram (gm)'),
        ('nos', 'Numbers (nos)'),
        ('mtr', 'Meter (mtr)'),
    ]

    assembly = models.ForeignKey(Assembly, on_delete=models.CASCADE, related_name='assembly_raw_materials')
    description = models.CharField(max_length=200, default="")
    production_quantity = models.IntegerField(default=1)
    production_weight = models.CharField(max_length=100, blank=True, null=True, default="",
                                        help_text="Optional production weight (text field, not used in calculations)")
    unit = models.CharField(max_length=10, choices=UNIT_CHOICES, default='kg')
    cost_per_unit = models.DecimalField(max_digits=18, decimal_places=8, default=0)

    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['created_at']
        verbose_name = 'Assembly Raw Material'
        verbose_name_plural = 'Assembly Raw Materials'

    def __str__(self):
        return f"{self.description} - {self.assembly.quote.name}"

    @property
    def total_cost(self):
        """Calculate total cost"""
        from decimal import Decimal
        if self.production_quantity > 0:
            total = Decimal(str(self.cost_per_unit)) / Decimal(str(self.production_quantity))
            return float(total)
        return 0

    def save(self, *args, **kwargs):
        """Override save to trigger assembly recalculation"""
        super().save(*args, **kwargs)
        self.assembly.save()


class ManufacturingPrintingCost(models.Model):
    """Manufacturing/Printing costs for assembly"""
    assembly = models.ForeignKey(Assembly, on_delete=models.CASCADE, related_name='manufacturing_printing_costs')

    process = models.CharField(max_length=200, default="", help_text="Process name/description")
    mc_tonnage = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                    verbose_name="M/C Tonnage", help_text="Machine tonnage")
    mc_rate_per_hour = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                          verbose_name="M/C Rate/Hr", help_text="Machine rate per hour")
    cycle_time = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                    help_text="Cycle time in seconds")

    # Calculated field
    per_cost = models.DecimalField(max_digits=18, decimal_places=8, default=0, editable=False,
                                  help_text="(Rate/Hr × Cycle Time) / 3600")

    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['created_at']
        verbose_name = 'Manufacturing/Printing Cost'
        verbose_name_plural = 'Manufacturing/Printing Costs'

    def __str__(self):
        return f"{self.process} - Assembly {self.assembly.id}"

    def save(self, *args, **kwargs):
        """Calculate per cost before saving"""
        from decimal import Decimal

        mc_rate_per_hour = Decimal(str(self.mc_rate_per_hour or 0))
        cycle_time = Decimal(str(self.cycle_time or 0))

        # Formula: (Rate/Hr × Cycle Time) / 3600
        self.per_cost = (mc_rate_per_hour * cycle_time) / Decimal('3600')

        super().save(*args, **kwargs)

        # Update parent assembly costs (avoid recursion issues)
        if self.assembly_id:
            self.assembly.calculate_costs()


class Packaging(models.Model):
    """Packaging for a quote"""
    PACKAGING_TYPE_CHOICES = [
        ('pp_box', 'PP Box'),
        ('pp_box_partition', 'PP Box by Partition'),
        ('cg_box', 'CG Box'),
        ('bin', 'Bin'),
        ('custom', 'Custom'),
    ]

    quote = models.ForeignKey(Quote, on_delete=models.CASCADE, related_name='packagings')
    packaging_type = models.ForeignKey(PackagingType, on_delete=models.SET_NULL,
                                      null=True, blank=True,
                                      related_name='packagings',
                                      help_text="Select packaging type")

    # Dimensions (editable)
    packaging_length = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                          help_text="Packaging length (mm)")
    packaging_breadth = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                           help_text="Packaging breadth (mm)")
    packaging_height = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                          help_text="Packaging height (mm)")
    polybag_length = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                        help_text="Polybag length")
    polybag_width = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                       help_text="Polybag width")

    # Cost details
    lifecycle = models.IntegerField(default=0, help_text="Lifecycle count")
    cost = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                              help_text="Packaging cost")
    maintenance_percentage = models.DecimalField(max_digits=13, decimal_places=8, default=0,
                                                help_text="Maintenance percentage")
    parts_per_polybag = models.IntegerField(default=0, help_text="Parts per polybag")

    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['created_at']
        verbose_name = 'Packaging'
        verbose_name_plural = 'Packagings'

    def __str__(self):
        type_display = self.packaging_type.name if self.packaging_type else "Custom"
        return f"{type_display} - {self.quote.name}"

    @property
    def maintenance_cost(self):
        """Calculate maintenance cost"""
        from decimal import Decimal
        return float(Decimal(str(self.cost)) * (Decimal(str(self.maintenance_percentage)) / Decimal('100')))

    @property
    def total_cost(self):
        """Calculate total packaging cost"""
        from decimal import Decimal
        return float(Decimal(str(self.cost)) + Decimal(str(self.maintenance_cost)))

    @property
    def cost_per_part(self):
        """Calculate packaging cost per part"""
        from decimal import Decimal

        if self.lifecycle > 0 and self.parts_per_polybag > 0:
            parts_per_pkg = Decimal(str(self.lifecycle)) * Decimal(str(self.parts_per_polybag))
            if parts_per_pkg > 0:
                return float(Decimal(str(self.total_cost)) / parts_per_pkg)
        return 0

    def save(self, *args, **kwargs):
        """Override save to auto-populate dimensions from packaging type if available"""
        if self.packaging_type and not self.pk:  # Only on creation
            if self.packaging_length == 0:
                self.packaging_length = self.packaging_type.default_length
            if self.packaging_breadth == 0:
                self.packaging_breadth = self.packaging_type.default_breadth
            if self.packaging_height == 0:
                self.packaging_height = self.packaging_type.default_height
            if self.polybag_length == 0:
                self.polybag_length = self.packaging_type.default_polybag_length
            if self.polybag_width == 0:
                self.polybag_width = self.packaging_type.default_polybag_width

        super().save(*args, **kwargs)

class Transport(models.Model):
    """Transport for a quote"""
    quote = models.ForeignKey(Quote, on_delete=models.CASCADE, related_name='transports')
    packaging = models.ForeignKey(Packaging, on_delete=models.SET_NULL, null=True, blank=True,
                                  related_name='transports', help_text="Related packaging")

    # Transport dimensions in feet
    transport_length = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                          help_text="Transport length (ft)")
    transport_breadth = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                           help_text="Transport breadth (ft)")
    transport_height = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                          help_text="Transport height (ft)")

    # Cost details
    trip_cost = models.DecimalField(max_digits=18, decimal_places=8, default=0,
                                   help_text="Cost per trip")
    parts_per_box = models.IntegerField(default=1, help_text="Number of parts per box")

    created_at = models.DateTimeField(default=timezone.now)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['created_at']
        verbose_name = 'Transport'
        verbose_name_plural = 'Transports'

    def __str__(self):
        return f"Transport - {self.quote.name}"

    @property
    def transport_length_mm(self):
        """Convert transport length from feet to mm (1 ft = 304.8 mm)"""
        from decimal import Decimal
        return float(Decimal(str(self.transport_length)) * Decimal('304.8'))

    @property
    def transport_breadth_mm(self):
        """Convert transport breadth from feet to mm (1 ft = 304.8 mm)"""
        from decimal import Decimal
        return float(Decimal(str(self.transport_breadth)) * Decimal('304.8'))

    @property
    def transport_height_mm(self):
        """Convert transport height from feet to mm (1 ft = 304.8 mm)"""
        from decimal import Decimal
        return float(Decimal(str(self.transport_height)) * Decimal('304.8'))

    @property
    def boxes_on_length(self):
        """Calculate boxes on length"""
        if self.packaging and self.transport_length_mm > 0 and self.packaging.packaging_length > 0:
            from decimal import Decimal
            return int(Decimal(str(self.transport_length_mm)) / Decimal(str(self.packaging.packaging_length)))
        return 0

    @property
    def boxes_on_breadth(self):
        """Calculate boxes on breadth"""
        if self.packaging and self.transport_breadth_mm > 0 and self.packaging.packaging_breadth > 0:
            from decimal import Decimal
            return int(Decimal(str(self.transport_breadth_mm)) / Decimal(str(self.packaging.packaging_breadth)))
        return 0

    @property
    def boxes_on_height(self):
        """Calculate boxes on height"""
        if self.packaging and self.transport_height_mm > 0 and self.packaging.packaging_height > 0:
            from decimal import Decimal
            return int(Decimal(str(self.transport_height_mm)) / Decimal(str(self.packaging.packaging_height)))
        return 0

    @property
    def total_boxes(self):
        """Calculate total boxes"""
        return self.boxes_on_length * self.boxes_on_breadth * self.boxes_on_height

    @property
    def total_parts_per_trip(self):
        """Calculate total parts per trip"""
        return self.total_boxes * self.parts_per_box

    @property
    def trip_cost_per_part(self):
        """Calculate trip cost per part"""
        if self.total_parts_per_trip > 0:
            from decimal import Decimal
            return float(Decimal(str(self.trip_cost)) / Decimal(str(self.total_parts_per_trip)))
        return 0


class QuoteTimeline(models.Model):
    """Timeline tracking for quote changes and activities"""
    ACTIVITY_TYPE_CHOICES = [
        ('quote_created', 'Quote Created'),
        ('quote_updated', 'Quote Updated'),
        ('raw_material_added', 'Raw Material Added'),
        ('raw_material_deleted', 'Raw Material Deleted'),
        ('moulding_machine_added', 'Moulding Machine Added'),
        ('moulding_machine_deleted', 'Moulding Machine Deleted'),
        ('assembly_added', 'Assembly Added'),
        ('assembly_deleted', 'Assembly Deleted'),
        ('assembly_rm_added', 'Assembly Raw Material Added'),
        ('assembly_rm_deleted', 'Assembly Raw Material Deleted'),
        ('manufacturing_cost_added', 'Manufacturing Cost Added'),
        ('manufacturing_cost_deleted', 'Manufacturing Cost Deleted'),
        ('packaging_added', 'Packaging Added'),
        ('packaging_deleted', 'Packaging Deleted'),
        ('transport_added', 'Transport Added'),
        ('transport_deleted', 'Transport Deleted'),
        ('section_completed', 'Section Completed'),
        ('manual_entry', 'Manual Entry'),
    ]

    quote = models.ForeignKey(Quote, on_delete=models.CASCADE, related_name='timeline_entries')
    activity_type = models.CharField(max_length=50, choices=ACTIVITY_TYPE_CHOICES)
    description = models.TextField(help_text="Description of the change/activity")
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='quote_timeline_entries')

    # For manual entries
    attachment = models.FileField(upload_to='timeline_attachments/', null=True, blank=True,
                                 help_text="Optional file attachment")

    created_at = models.DateTimeField(default=timezone.now)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Quote Timeline Entry'
        verbose_name_plural = 'Quote Timeline Entries'

    def __str__(self):
        return f"{self.get_activity_type_display()} - {self.quote.name} - {self.created_at}"

    @staticmethod
    def add_entry(quote, activity_type, description, user):
        """Helper method to add timeline entry"""
        QuoteTimeline.objects.create(
            quote=quote,
            activity_type=activity_type,
            description=description,
            user=user
        )
