
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from .models import (
    Project, Quote, CustomerGroup, MaterialGroup,
    AssemblyType, PackagingType, RawMaterial, MouldingMachineDetail,
    Assembly, AssemblyRawMaterial, ManufacturingPrintingCost, Packaging, Transport,
    QuoteTimeline, MaterialType, MouldingMachineType
)
from django.contrib.auth.models import User
from django.contrib.auth.decorators import user_passes_test
from django.http import HttpResponse
from .excel_utils import ExcelTemplateGenerator, ExcelParser


# Check if user is superuser
def superuser_required(user):
    return user.is_superuser


@login_required
def home(request):
    return render(request, 'core/home.html')


@login_required
def projects(request):
    """List all projects"""
    all_projects = Project.objects.filter(is_active=True)
    return render(request, 'core/projects.html', {'projects': all_projects})


@login_required
def project_create(request):
    """Create a new project"""
    if request.method == 'POST':
        name = request.POST.get('name')
        description = request.POST.get('description')

        if name:
            project = Project.objects.create(
                name=name,
                description=description,
                created_by=request.user
            )
            messages.success(request, f'Project "{project.name}" created successfully!')
            return redirect('project_detail', project_id=project.id)
        else:
            messages.error(request, 'Project name is required.')

    return render(request, 'core/project_create.html')


@login_required
def project_detail(request, project_id):
    """View individual project with its quotes"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quotes = project.quotes.all()

    context = {
        'project': project,
        'quotes': quotes,
    }
    return render(request, 'core/project_detail.html', context)


@login_required
def quote_create(request, project_id):
    """Create a new quote with quote definition"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    customer_groups = CustomerGroup.objects.filter(is_active=True)

    if request.method == 'POST':
        name = request.POST.get('name')
        client_group_id = request.POST.get('client_group')
        client_name = request.POST.get('client_name')
        notes = request.POST.get('notes', '')
        handling_charge = request.POST.get('handling_charge', 0)
        profit_percentage = request.POST.get('profit_percentage', 0)
        sap_number = request.POST.get('sap_number', '')
        part_number = request.POST.get('part_number')
        part_name = request.POST.get('part_name')
        amendment_number = request.POST.get('amendment_number', '')
        description = request.POST.get('description', '')
        quantity = int(request.POST.get('quantity', 1))

        if name and client_group_id and client_name and part_number and part_name:
            quote = Quote.objects.create(
                project=project,
                name=name,
                client_group_id=client_group_id,
                client_name=client_name,
                notes=notes,
                handling_charge=float(handling_charge),
                profit_percentage=float(profit_percentage),
                sap_number=sap_number,
                part_number=part_number,
                part_name=part_name,
                amendment_number=amendment_number,
                description=description,
                quantity=quantity,
                created_by=request.user,
                quote_definition_complete=True
            )

            # Add timeline entry
            QuoteTimeline.add_entry(
                quote=quote,
                activity_type='quote_created',
                description=f'Quote "{quote.name}" created for client {client_name}',
                user=request.user
            )

            messages.success(request, f'Quote "{quote.name}" created successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        else:
            messages.error(request, 'Please fill in all required fields including Customer Group.')

    context = {
        'project': project,
        'customer_groups': customer_groups,
    }
    return render(request, 'core/quote_create.html', context)


@login_required
def quote_detail(request, project_id, quote_id):
    """View individual quote with all sections"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    raw_materials = quote.raw_materials.all()
    moulding_machines = quote.moulding_machines.all()
    assemblies = quote.assemblies.all()
    packagings = quote.packagings.all()
    transports = quote.transports.all()
    timeline_entries = quote.timeline_entries.all()

    context = {
        'project': project,
        'quote': quote,
        'raw_materials': raw_materials,
        'moulding_machines': moulding_machines,
        'assemblies': assemblies,
        'packagings': packagings,
        'transports': transports,
        'timeline_entries': timeline_entries,
        'total_rm_cost': quote.get_total_raw_material_cost(),
        'total_conversion_cost': quote.get_total_conversion_cost(),
        'total_assembly_cost': quote.get_total_assembly_cost(),
        'total_packaging_cost': quote.get_total_packaging_cost(),
        'total_transport_cost': quote.get_total_transport_cost(),
    }
    return render(request, 'core/quote_detail.html', context)


@login_required
def quote_definition_edit(request, project_id, quote_id):
    """Edit quote definition"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    customer_groups = CustomerGroup.objects.filter(is_active=True)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        quote.name = request.POST.get('name')
        client_group_id = request.POST.get('client_group')
        if not client_group_id:
            messages.error(request, 'Customer Group is required.')
            context = {
                'project': project,
                'quote': quote,
                'customer_groups': customer_groups,
            }
            return render(request, 'core/quote_definition_edit.html', context)

        quote.client_group_id = client_group_id
        quote.client_name = request.POST.get('client_name')
        quote.notes = request.POST.get('notes')
        quote.handling_charge = float(request.POST.get('handling_charge', 0))
        quote.profit_percentage = float(request.POST.get('profit_percentage', 0))
        quote.sap_number = request.POST.get('sap_number')
        quote.part_number = request.POST.get('part_number')
        quote.part_name = request.POST.get('part_name')
        quote.amendment_number = request.POST.get('amendment_number')
        quote.description = request.POST.get('description')
        quote.quantity = request.POST.get('quantity', 1)
        quote.quote_definition_complete = True

        quote.save()
        quote.increment_version(request.user, 'Quote definition updated')

        messages.success(request, 'Quote definition updated successfully!')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    context = {
        'project': project,
        'quote': quote,
        'customer_groups': customer_groups,
    }
    return render(request, 'core/quote_definition_edit.html', context)


# ============ RAW MATERIALS ============


@login_required
def raw_material_add(request, project_id, quote_id):
    """Add raw material to quote"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Get material types for the quote's customer group
    material_types = MaterialType.objects.filter(
        customer_group=quote.client_group,
        is_active=True
    ) if quote.client_group else []

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            material_type_id = request.POST.get('material_type')
            frozen_rate_str = request.POST.get('frozen_rate')
            frozen_rate = float(frozen_rate_str) if frozen_rate_str else None

            raw_material = RawMaterial(
                quote=quote,
                material_type_id=material_type_id if material_type_id else None,
                material_name=request.POST.get('material_name'),
                grade=request.POST.get('grade'),
                rm_code=request.POST.get('rm_code'),
                unit_of_measurement=request.POST.get('unit_of_measurement', 'kg'),
                frozen_rate=frozen_rate,
                rm_rate=float(request.POST.get('rm_rate')),
                part_weight=float(request.POST.get('part_weight')),
                runner_weight=float(request.POST.get('runner_weight')),
                process_losses=float(request.POST.get('process_losses', 0)),
                purging_loss_cost=float(request.POST.get('purging_loss_cost', 0)),
                icc_percentage=float(request.POST.get('icc_percentage', 0)),
                rejection_percentage=float(request.POST.get('rejection_percentage', 0)),
                overhead_percentage=float(request.POST.get('overhead_percentage', 0)),
                maintenance_percentage=float(request.POST.get('maintenance_percentage', 0)),
                profit_percentage=float(request.POST.get('profit_percentage', 0)),
            )
            raw_material.save()

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

            messages.success(request, f'Raw material "{raw_material.material_name}" added successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error adding raw material: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'material_types': material_types,
        'unit_choices': RawMaterial.UNIT_CHOICES,
    }
    return render(request, 'core/raw_material_add.html', context)


@login_required
def raw_material_delete(request, project_id, quote_id, rm_id):
    """Delete a raw material"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    raw_material = get_object_or_404(RawMaterial, id=rm_id, quote=quote)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    material_name = raw_material.material_name
    raw_material.delete()

    # Add timeline entry
    QuoteTimeline.add_entry(
        quote=quote,
        activity_type='raw_material_deleted',
        description=f'Raw material "{material_name}" deleted',
        user=request.user
    )
    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

    messages.success(request, f'Raw material "{material_name}" deleted successfully!')
    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


@login_required
def raw_material_complete(request, project_id, quote_id):
    """Mark raw material section as complete"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if quote.raw_materials.count() > 0:
        quote.raw_material_complete = True
        quote.save()

        # Add timeline entry
        QuoteTimeline.add_entry(
            quote=quote,
            activity_type='section_completed',
            description='Raw material section marked as complete',
            user=request.user
        )
        # Increment version and add timeline entry
        quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')


        messages.success(request, 'Raw material section marked as complete!')
    else:
        messages.error(request, 'Please add at least one raw material before marking as complete.')

    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


# ============ MOULDING MACHINES ============

@login_required
def moulding_machine_add(request, project_id, quote_id):
    """Add moulding machine to quote"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Get moulding machine types for the quote's customer group
    moulding_machine_types = MouldingMachineType.objects.filter(
        customer_group=quote.client_group,
        is_active=True
    ) if quote.client_group else []

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            moulding_machine_type_id = request.POST.get('moulding_machine_type')

            machine = MouldingMachineDetail.objects.create(
                quote=quote,
                moulding_machine_type_id=moulding_machine_type_id if moulding_machine_type_id else None,
                cavity=int(request.POST.get('cavity', 1)),
                machine_tonnage=float(request.POST.get('machine_tonnage', 0)),
                cycle_time=float(request.POST.get('cycle_time', 0)),
                efficiency=float(request.POST.get('efficiency', 0)),
                shift_rate=float(request.POST.get('shift_rate', 0)),
                shift_rate_for_mtc=float(request.POST.get('shift_rate_for_mtc', 0)),
                mtc_cost=float(request.POST.get('mtc_cost', 0)),
                rejection_percentage=float(request.POST.get('rejection_percentage', 0)),
                overhead_percentage=float(request.POST.get('overhead_percentage', 0)),
                maintenance_percentage=float(request.POST.get('maintenance_percentage', 0)),
                profit_percentage=float(request.POST.get('profit_percentage', 0)),
            )

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Moulding machine with {machine.cavity} cavity added')

            messages.success(request, 'Moulding machine added successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error adding moulding machine: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'moulding_machine_types': moulding_machine_types,
    }
    return render(request, 'core/moulding_machine_add.html', context)


@login_required
def moulding_machine_delete(request, project_id, quote_id, mm_id):
    """Delete a moulding machine detail"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    moulding_machine = get_object_or_404(MouldingMachineDetail, id=mm_id, quote=quote)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    cavity = moulding_machine.cavity
    moulding_machine.delete()

    # Add timeline entry
    QuoteTimeline.add_entry(
        quote=quote,
        activity_type='moulding_machine_deleted',
        description=f'Moulding machine with {cavity} cavities deleted',
        user=request.user
    )
    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

    messages.success(request, 'Moulding machine detail deleted successfully!')
    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


@login_required
def moulding_machine_complete(request, project_id, quote_id):
    """Mark moulding machine section as complete"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if quote.moulding_machines.count() > 0:
        quote.moulding_machine_complete = True
        quote.save()

        # Add timeline entry
        QuoteTimeline.add_entry(
            quote=quote,
            activity_type='section_completed',
            description='Moulding machine section marked as complete',
            user=request.user
        )
        # Increment version and add timeline entry
        quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

        messages.success(request, 'Moulding machine section marked as complete!')
    else:
        messages.error(request, 'Please add at least one moulding machine before marking as complete.')

    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


# ============ ASSEMBLY ============

@login_required
def assembly_add(request, project_id, quote_id):
    """Add assembly to quote"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly_types = AssemblyType.objects.filter(is_active=True)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            assembly_type = request.POST.get('assembly_type', 'manual')

            assembly = Assembly(
                quote=quote,
                assembly_type=assembly_type,
                manual_cost=float(request.POST.get('manual_cost', 0)) if assembly_type == 'manual' else 0,
                assembly_type_config_id=request.POST.get('assembly_type_config') or None if assembly_type == 'automated' else None,
                other_cost=float(request.POST.get('other_cost', 0)),
                profit_percentage=float(request.POST.get('profit_percentage', 0)),
                rejection_percentage=float(request.POST.get('rejection_percentage', 0)),
                inspection_handling_percentage=float(request.POST.get('inspection_handling_percentage', 0)),
            )
            assembly.save()
            assembly.calculate_costs()

            # Add timeline entry
            QuoteTimeline.add_entry(
                quote=quote,
                activity_type='assembly_added',
                description=f'{assembly.get_assembly_type_display()} assembly added',
                user=request.user
            )
            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

            messages.success(request, 'Assembly added successfully!')
            return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error adding assembly: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'assembly_types': assembly_types,
    }
    return render(request, 'core/assembly_add.html', context)


@login_required
def assembly_detail(request, project_id, quote_id, assembly_id):
    """View assembly detail with raw materials and manufacturing costs"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly = get_object_or_404(Assembly, id=assembly_id, quote=quote)

    context = {
        'project': project,
        'quote': quote,
        'assembly': assembly,
        'assembly_raw_materials': assembly.assembly_raw_materials.all(),
        'manufacturing_printing_costs': assembly.manufacturing_printing_costs.all(),
    }
    return render(request, 'core/assembly_detail.html', context)


@login_required
def assembly_delete(request, project_id, quote_id, assembly_id):
    """Delete an assembly"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly = get_object_or_404(Assembly, id=assembly_id, quote=quote)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    assembly_type = assembly.get_assembly_type_display()
    assembly.delete()

    # Add timeline entry
    QuoteTimeline.add_entry(
        quote=quote,
        activity_type='assembly_deleted',
        description=f'{assembly_type} assembly deleted',
        user=request.user
    )
    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

    messages.success(request, 'Assembly deleted successfully!')
    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


@login_required
def assembly_raw_material_add(request, project_id, quote_id, assembly_id):
    """Add assembly raw material"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly = get_object_or_404(Assembly, id=assembly_id, quote=quote)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            arm = AssemblyRawMaterial(
                assembly=assembly,
                description=request.POST.get('description'),
                production_quantity=int(request.POST.get('production_quantity', 1)),
                production_weight=float(request.POST.get('production_weight', 0)),
                unit=request.POST.get('unit', 'kg'),
                cost_per_unit=float(request.POST.get('cost_per_unit', 0)),
            )
            arm.save()

            # Add timeline entry
            QuoteTimeline.add_entry(
                quote=quote,
                activity_type='assembly_rm_added',
                description=f'Assembly raw material "{arm.description}" added to assembly #{assembly.id}',
                user=request.user
            )
            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

            messages.success(request, 'Assembly raw material added successfully!')
            return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)
        except ValueError:
            messages.error(request, 'Invalid numeric value provided.')
        except Exception as e:
            messages.error(request, f'Error: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'assembly': assembly,
        'unit_choices': AssemblyRawMaterial.UNIT_CHOICES,
    }
    return render(request, 'core/assembly_raw_material_add.html', context)


@login_required
def assembly_raw_material_delete(request, project_id, quote_id, assembly_id, arm_id):
    """Delete assembly raw material"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly = get_object_or_404(Assembly, id=assembly_id, quote=quote)
    arm = get_object_or_404(AssemblyRawMaterial, id=arm_id, assembly=assembly)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    description = arm.description
    arm.delete()

    # Add timeline entry
    QuoteTimeline.add_entry(
        quote=quote,
        activity_type='assembly_rm_deleted',
        description=f'Assembly raw material "{description}" deleted from assembly #{assembly.id}',
        user=request.user
    )
    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

    messages.success(request, 'Assembly raw material deleted successfully!')
    return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)


@login_required
def manufacturing_printing_cost_add(request, project_id, quote_id, assembly_id):
    """Add manufacturing/printing cost"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly = get_object_or_404(Assembly, id=assembly_id, quote=quote)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            mpc = ManufacturingPrintingCost(
                assembly=assembly,
                process=request.POST.get('process'),
                mc_tonnage=float(request.POST.get('mc_tonnage', 0)),
                mc_rate_per_hour=float(request.POST.get('mc_rate_per_hour', 0)),
                cycle_time=float(request.POST.get('cycle_time', 0)),
            )
            mpc.save()

            # Add timeline entry
            QuoteTimeline.add_entry(
                quote=quote,
                activity_type='manufacturing_cost_added',
                description=f'Manufacturing cost "{mpc.process}" added to assembly #{assembly.id}',
                user=request.user
            )
            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

            messages.success(request, 'Manufacturing/printing cost added successfully!')
            return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)
        except ValueError:
            messages.error(request, 'Invalid numeric value provided.')
        except Exception as e:
            messages.error(request, f'Error: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'assembly': assembly,
    }
    return render(request, 'core/manufacturing_printing_cost_add.html', context)


@login_required
def manufacturing_printing_cost_delete(request, project_id, quote_id, assembly_id, mpc_id):
    """Delete manufacturing/printing cost"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly = get_object_or_404(Assembly, id=assembly_id, quote=quote)
    mpc = get_object_or_404(ManufacturingPrintingCost, id=mpc_id, assembly=assembly)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    process = mpc.process
    mpc.delete()

    # Add timeline entry
    QuoteTimeline.add_entry(
        quote=quote,
        activity_type='manufacturing_cost_deleted',
        description=f'Manufacturing cost "{process}" deleted from assembly #{assembly.id}',
        user=request.user
    )
    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

    messages.success(request, 'Manufacturing/printing cost deleted successfully!')
    return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)


@login_required
def assembly_complete(request, project_id, quote_id):
    """Mark assembly section as complete"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if quote.assemblies.count() > 0:
        quote.assembly_complete = True
        quote.save()

        # Add timeline entry
        QuoteTimeline.add_entry(
            quote=quote,
            activity_type='section_completed',
            description='Assembly section marked as complete',
            user=request.user
        )
        # Increment version and add timeline entry
        quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

        messages.success(request, 'Assembly section marked as complete!')
    else:
        messages.error(request, 'Please add at least one assembly before marking as complete.')

    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


# ============ PACKAGING ============

@login_required
def packaging_add(request, project_id, quote_id):
    """Add packaging to quote"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            packaging = Packaging(
                quote=quote,
                packaging_type=request.POST.get('packaging_type', 'pp_box'),
                lifecycle=int(request.POST.get('lifecycle', 0)),
                cost=float(request.POST.get('cost', 0)),
                maintenance_percentage=float(request.POST.get('maintenance_percentage', 0)),
                part_per_polybag=int(request.POST.get('part_per_polybag', 1)),
            )
            packaging.save()

            # Add timeline entry
            QuoteTimeline.add_entry(
                quote=quote,
                activity_type='packaging_added',
                description=f'Packaging "{packaging.get_packaging_type_display()}" added',
                user=request.user
            )
            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

            messages.success(request, f'Packaging "{packaging.get_packaging_type_display()}" added successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except ValueError:
            messages.error(request, 'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error adding packaging: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'packaging_types': Packaging.PACKAGING_TYPE_CHOICES,
    }
    return render(request, 'core/packaging_add.html', context)


@login_required
def packaging_delete(request, project_id, quote_id, packaging_id):
    """Delete a packaging"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    packaging = get_object_or_404(Packaging, id=packaging_id, quote=quote)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    packaging_type = packaging.get_packaging_type_display()
    packaging.delete()

    # Add timeline entry
    QuoteTimeline.add_entry(
        quote=quote,
        activity_type='packaging_deleted',
        description=f'Packaging "{packaging_type}" deleted',
        user=request.user
    )
    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

    messages.success(request, f'Packaging "{packaging_type}" deleted successfully!')
    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


@login_required
def packaging_complete(request, project_id, quote_id):
    """Mark packaging section as complete"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if quote.packagings.count() > 0:
        quote.packaging_complete = True
        quote.save()

        # Add timeline entry
        QuoteTimeline.add_entry(
            quote=quote,
            activity_type='section_completed',
            description='Packaging section marked as complete',
            user=request.user
        )
        # Increment version and add timeline entry
        quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

        messages.success(request, 'Packaging section marked as complete!')
    else:
        messages.error(request, 'Please add at least one packaging before marking as complete.')

    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


# ============ TRANSPORT ============

@login_required
def transport_add(request, project_id, quote_id):
    """Add transport to quote"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    packagings = quote.packagings.all()

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if not packagings:
        messages.error(request, 'Please add packaging first before adding transport.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            transport = Transport(
                quote=quote,
                packaging_id=request.POST.get('packaging'),
                transport_length=float(request.POST.get('transport_length', 0)),
                transport_breadth=float(request.POST.get('transport_breadth', 0)),
                transport_height=float(request.POST.get('transport_height', 0)),
                trip_cost=float(request.POST.get('trip_cost', 0)),
                parts_per_box=int(request.POST.get('parts_per_box', 1)),
            )
            transport.save()

            # Add timeline entry
            QuoteTimeline.add_entry(
                quote=quote,
                activity_type='transport_added',
                description=f'Transport for {transport.packaging.get_packaging_type_display()} added',
                user=request.user
            )
            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

            messages.success(request, 'Transport added successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except ValueError:
            messages.error(request, 'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error adding transport: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'packagings': packagings,
    }
    return render(request, 'core/transport_add.html', context)


@login_required
def transport_delete(request, project_id, quote_id, transport_id):
    """Delete a transport"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    transport = get_object_or_404(Transport, id=transport_id, quote=quote)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    packaging_type = transport.packaging.get_packaging_type_display()
    transport.delete()

    # Add timeline entry
    QuoteTimeline.add_entry(
        quote=quote,
        activity_type='transport_deleted',
        description=f'Transport for {packaging_type} deleted',
        user=request.user
    )
    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

    messages.success(request, 'Transport deleted successfully!')
    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


@login_required
def transport_complete(request, project_id, quote_id):
    """Mark transport section as complete"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if quote.transports.count() > 0:
        quote.transport_complete = True
        quote.save()

        # Add timeline entry
        QuoteTimeline.add_entry(
            quote=quote,
            activity_type='section_completed',
            description='Transport section marked as complete',
            user=request.user
        )
        # Increment version and add timeline entry
        quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

        messages.success(request, 'Transport section marked as complete!')
    else:
        messages.error(request, 'Please add at least one transport before marking as complete.')

    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


# ============ TIMELINE ============

@login_required
def timeline_add_manual(request, project_id, quote_id):
    """Add manual timeline entry"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        description = request.POST.get('description')
        attachment = request.FILES.get('attachment')

        if description:
            timeline_entry = QuoteTimeline.objects.create(
                quote=quote,
                activity_type='manual_entry',
                description=description,
                user=request.user,
                attachment=attachment
            )
            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added')

            messages.success(request, 'Timeline entry added successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        else:
            messages.error(request, 'Description is required.')

    context = {
        'project': project,
        'quote': quote,
    }
    return render(request, 'core/timeline_add_manual.html', context)


# ============ CONFIG ============

@login_required
def config(request):
    """Configuration page with all config types"""
    customer_groups = CustomerGroup.objects.filter(is_active=True)

    # Get selected customer group - default to first one if none selected
    selected_customer_group_id = request.GET.get('customer_group')
    selected_customer_group = None

    if selected_customer_group_id:
        selected_customer_group = get_object_or_404(CustomerGroup, id=selected_customer_group_id)
    elif customer_groups.exists():
        # Default to first customer group
        selected_customer_group = customer_groups.first()

    # Get all config items for selected customer group
    material_types = []
    moulding_machine_types = []
    assembly_types = []
    packaging_types = []

    if selected_customer_group:
        material_types = MaterialType.objects.filter(customer_group=selected_customer_group, is_active=True)
        moulding_machine_types = MouldingMachineType.objects.filter(customer_group=selected_customer_group, is_active=True)
        assembly_types = AssemblyType.objects.filter(customer_group=selected_customer_group, is_active=True)
        packaging_types = PackagingType.objects.filter(customer_group=selected_customer_group, is_active=True)

    context = {
        'customer_groups': customer_groups,
        'selected_customer_group': selected_customer_group,
        'material_types': material_types,
        'moulding_machine_types': moulding_machine_types,
        'assembly_types': assembly_types,
        'packaging_types': packaging_types,
    }
    return render(request, 'core/config.html', context)


@login_required
def config_create(request, config_type):
    """Create a new configuration item"""
    # Get customer group
    customer_group_id = request.GET.get('customer_group')
    if not customer_group_id:
        messages.error(request, 'Customer group is required.')
        return redirect('config')

    customer_group = get_object_or_404(CustomerGroup, id=customer_group_id)

    # Map config_type to model (removed material-group)
    model_map = {
        'assembly-type': AssemblyType,
        'packaging-type': PackagingType,
    }

    title_map = {
        'assembly-type': 'Assembly Type',
        'packaging-type': 'Packaging Type',
    }

    if config_type not in model_map:
        messages.error(request, 'Invalid configuration type.')
        return redirect('config')

    model = model_map[config_type]
    title = title_map[config_type]

    if request.method == 'POST':
        name = request.POST.get('name')
        value = request.POST.get('value')
        description = request.POST.get('description')

        if name and value:
            try:
                item = model.objects.create(
                    customer_group=customer_group,
                    name=name,
                    value=value,
                    description=description,
                    created_by=request.user
                )
                messages.success(request, f'{title} "{item.name}" created successfully!')
                return redirect(f'/config/?customer_group={customer_group_id}')
            except Exception as e:
                messages.error(request, f'Error creating {title}: {str(e)}')
        else:
            messages.error(request, 'Name and Value are required.')

    context = {
        'config_type': config_type,
        'title': title,
        'customer_group': customer_group,
    }
    return render(request, 'core/config_create.html', context)


@login_required
def config_delete(request, config_type, item_id):
    """Soft delete a configuration item"""
    model_map = {
        'assembly-type': AssemblyType,
        'packaging-type': PackagingType,
    }

    title_map = {
        'assembly-type': 'Assembly Type',
        'packaging-type': 'Packaging Type',
    }

    if config_type not in model_map:
        messages.error(request, 'Invalid configuration type.')
        return redirect('config')

    model = model_map[config_type]
    title = title_map[config_type]

    item = get_object_or_404(model, id=item_id)
    customer_group_id = item.customer_group.id
    item.is_active = False
    item.save()

    messages.success(request, f'{title} "{item.name}" deleted successfully!')
    return redirect(f'/config/?customer_group={customer_group_id}')


@login_required
def quote_summary(request, project_id, quote_id):
    """View quote summary with all costs"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Get all data
    raw_materials = quote.raw_materials.all()
    moulding_machines = quote.moulding_machines.all()
    assemblies = quote.assemblies.all()
    packagings = quote.packagings.all()
    transports = quote.transports.all()

    # Calculate costs
    total_rm_cost = quote.get_total_raw_material_cost()
    total_conversion_cost = quote.get_total_conversion_cost()
    total_assembly_cost = quote.get_total_assembly_cost()
    total_packaging_cost = quote.get_total_packaging_cost()
    total_transport_cost = quote.get_total_transport_cost()

    # Calculate totals including profit and handling charge
    base_cost = quote.get_base_cost()
    profit_amount = quote.get_profit_amount()
    grand_total = quote.get_grand_total()

    context = {
        'project': project,
        'quote': quote,
        'raw_materials': raw_materials,
        'moulding_machines': moulding_machines,
        'assemblies': assemblies,
        'packagings': packagings,
        'transports': transports,
        'total_rm_cost': total_rm_cost,
        'total_conversion_cost': total_conversion_cost,
        'total_assembly_cost': total_assembly_cost,
        'total_packaging_cost': total_packaging_cost,
        'total_transport_cost': total_transport_cost,
        'base_cost': base_cost,
        'profit_amount': profit_amount,
        'grand_total': grand_total,
        'raw_materials_count': raw_materials.count(),
        'moulding_machines_count': moulding_machines.count(),
        'assemblies_count': assemblies.count(),
        'packagings_count': packagings.count(),
        'transports_count': transports.count(),
    }
    return render(request, 'core/quote_summary.html', context)

@login_required
@user_passes_test(superuser_required)
def admin_dashboard(request):
    """Admin dashboard for superuser only"""
    users = User.objects.all().order_by('-date_joined')
    total_users = users.count()
    active_users = users.filter(is_active=True).count()
    staff_users = users.filter(is_staff=True).count()

    context = {
        'users': users,
        'total_users': total_users,
        'active_users': active_users,
        'staff_users': staff_users,
    }
    return render(request, 'core/admin_dashboard.html', context)


@login_required
@user_passes_test(superuser_required)
def admin_user_create(request):
    """Create a new user"""
    if request.method == 'POST':
        username = request.POST.get('username')
        email = request.POST.get('email')
        password = request.POST.get('password')
        first_name = request.POST.get('first_name', '')
        last_name = request.POST.get('last_name', '')
        is_staff = request.POST.get('is_staff') == 'on'
        is_active = request.POST.get('is_active') == 'on'

        if username and password:
            # Check if username already exists
            if User.objects.filter(username=username).exists():
                messages.error(request, f'Username "{username}" already exists.')
            else:
                try:
                    user = User.objects.create_user(
                        username=username,
                        email=email,
                        password=password,
                        first_name=first_name,
                        last_name=last_name,
                        is_staff=is_staff,
                        is_active=is_active
                    )
                    messages.success(request, f'User "{user.username}" created successfully!')
                    return redirect('admin_dashboard')
                except Exception as e:
                    messages.error(request, f'Error creating user: {str(e)}')
        else:
            messages.error(request, 'Username and password are required.')

    return render(request, 'core/admin_user_create.html')


@login_required
@user_passes_test(superuser_required)
def admin_user_edit(request, user_id):
    """Edit an existing user"""
    user_to_edit = get_object_or_404(User, id=user_id)

    if request.method == 'POST':
        user_to_edit.username = request.POST.get('username')
        user_to_edit.email = request.POST.get('email')
        user_to_edit.first_name = request.POST.get('first_name', '')
        user_to_edit.last_name = request.POST.get('last_name', '')
        user_to_edit.is_staff = request.POST.get('is_staff') == 'on'
        user_to_edit.is_active = request.POST.get('is_active') == 'on'

        # Update password only if provided
        new_password = request.POST.get('password')
        if new_password:
            user_to_edit.set_password(new_password)

        try:
            user_to_edit.save()
            messages.success(request, f'User "{user_to_edit.username}" updated successfully!')
            return redirect('admin_dashboard')
        except Exception as e:
            messages.error(request, f'Error updating user: {str(e)}')

    context = {
        'user_to_edit': user_to_edit,
    }
    return render(request, 'core/admin_user_edit.html', context)


@login_required
@user_passes_test(superuser_required)
def admin_user_delete(request, user_id):
    """Delete a user"""
    user_to_delete = get_object_or_404(User, id=user_id)

    # Prevent deleting yourself
    if user_to_delete == request.user:
        messages.error(request, 'You cannot delete your own account.')
        return redirect('admin_dashboard')

    username = user_to_delete.username
    user_to_delete.delete()
    messages.success(request, f'User "{username}" deleted successfully!')
    return redirect('admin_dashboard')


@login_required
@user_passes_test(superuser_required)
def admin_user_toggle_active(request, user_id):
    """Toggle user active status"""
    user_to_toggle = get_object_or_404(User, id=user_id)

    # Prevent deactivating yourself
    if user_to_toggle == request.user:
        messages.error(request, 'You cannot deactivate your own account.')
        return redirect('admin_dashboard')

    user_to_toggle.is_active = not user_to_toggle.is_active
    user_to_toggle.save()

    status = "activated" if user_to_toggle.is_active else "deactivated"
    messages.success(request, f'User "{user_to_toggle.username}" {status} successfully!')
    return redirect('admin_dashboard')


@login_required
def quote_mark_completed(request, project_id, quote_id):
    """Mark quote as completed"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    if quote.mark_completed(request.user):
        messages.success(request, f'Quote marked as completed! Version updated to {quote.get_version()}')
    else:
        messages.error(request, 'Quote cannot be marked as completed.')

    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


@login_required
def quote_reopen(request, project_id, quote_id):
    """Reopen completed quote for editing"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    if quote.reopen_quote(request.user):
        messages.success(request, f'Quote reopened for editing! Version updated to {quote.get_version()}')
    else:
        messages.error(request, 'Quote cannot be reopened.')

    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


@login_required
def quote_discard(request, project_id, quote_id):
    """Discard a completed quote"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    if quote.discard_quote(request.user):
        messages.success(request, 'Quote has been discarded.')
    else:
        messages.error(request, 'Only completed quotes can be discarded.')

    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


@login_required
def material_type_create(request, customer_group_id):
    """Create a material type for a customer group"""
    customer_group = get_object_or_404(CustomerGroup, id=customer_group_id)

    if request.method == 'POST':
        raw_material_name = request.POST.get('raw_material_name')
        raw_material_grade = request.POST.get('raw_material_grade')
        raw_material_code = request.POST.get('raw_material_code')
        raw_material_rate = request.POST.get('raw_material_rate')

        if raw_material_name and raw_material_grade and raw_material_code and raw_material_rate:
            try:
                material_type = MaterialType.objects.create(
                    customer_group=customer_group,
                    raw_material_name=raw_material_name,
                    raw_material_grade=raw_material_grade,
                    raw_material_code=raw_material_code,
                    raw_material_rate=float(raw_material_rate),
                    created_by=request.user
                )
                messages.success(request, f'Material type "{material_type.raw_material_name}" created successfully!')
                return redirect(f'/config/?customer_group={customer_group_id}#material-types')
            except Exception as e:
                messages.error(request, f'Error creating material type: {str(e)}')
        else:
            messages.error(request, 'All fields are required.')

    context = {
        'customer_group': customer_group,
    }
    return render(request, 'core/material_type_create.html', context)

@login_required
def material_type_delete(request, material_type_id):
    """Delete a material type"""
    material_type = get_object_or_404(MaterialType, id=material_type_id)
    customer_group_id = material_type.customer_group.id

    material_type.is_active = False
    material_type.save()

    messages.success(request, f'Material type "{material_type.raw_material_name}" deleted successfully!')
    return redirect(f'/config/?customer_group={customer_group_id}#material-types')


@login_required
def moulding_machine_type_create(request, customer_group_id):
    """Create a moulding machine type for a customer group"""
    customer_group = get_object_or_404(CustomerGroup, id=customer_group_id)

    if request.method == 'POST':
        name = request.POST.get('name')
        shift_rate = request.POST.get('shift_rate')
        shift_rate_for_mtc = request.POST.get('shift_rate_for_mtc')
        mtc_cost = request.POST.get('mtc_cost')

        if name and shift_rate and shift_rate_for_mtc and mtc_cost:
            try:
                machine_type = MouldingMachineType.objects.create(
                    customer_group=customer_group,
                    name=name,
                    shift_rate=float(shift_rate),
                    shift_rate_for_mtc=float(shift_rate_for_mtc),
                    mtc_cost=float(mtc_cost),
                    created_by=request.user
                )
                messages.success(request, f'Moulding machine type "{machine_type.name}" created successfully!')
                return redirect(f'/config/?customer_group={customer_group_id}#moulding-machine-types')
            except Exception as e:
                messages.error(request, f'Error creating moulding machine type: {str(e)}')
        else:
            messages.error(request, 'All fields are required.')

    context = {
        'customer_group': customer_group,
    }
    return render(request, 'core/moulding_machine_type_create.html', context)

@login_required
def moulding_machine_type_delete(request, machine_type_id):
    """Delete a moulding machine type"""
    machine_type = get_object_or_404(MouldingMachineType, id=machine_type_id)
    customer_group_id = machine_type.customer_group.id

    machine_type.is_active = False
    machine_type.save()

    messages.success(request, 'Moulding machine type deleted successfully!')
    return redirect(f'/config/?customer_group={customer_group_id}#moulding-machine-types')


@login_required
def customer_group_create(request):
    """Create a new customer group"""
    if request.method == 'POST':
        name = request.POST.get('name')
        value = request.POST.get('value')
        description = request.POST.get('description')

        if name and value:
            try:
                customer_group = CustomerGroup.objects.create(
                    name=name,
                    value=value,
                    description=description,
                    created_by=request.user
                )
                messages.success(request, f'Customer group "{customer_group.name}" created successfully!')
                return redirect(f'/config/?customer_group={customer_group.id}')
            except Exception as e:
                messages.error(request, f'Error creating customer group: {str(e)}')
        else:
            messages.error(request, 'Name and Value are required.')

    return render(request, 'core/customer_group_create.html')


@login_required
def customer_group_delete(request, customer_group_id):
    """Delete a customer group"""
    customer_group = get_object_or_404(CustomerGroup, id=customer_group_id)

    # Check if customer group has any associated data
    if (customer_group.material_types.exists() or
        customer_group.moulding_machine_types.exists() or
        customer_group.material_groups.exists() or
        customer_group.assembly_types.exists() or
        customer_group.packaging_types.exists()):
        messages.error(request, 'Cannot delete customer group with associated configuration items.')
        return redirect('config')

    customer_group.is_active = False
    customer_group.save()

    messages.success(request, f'Customer group "{customer_group.name}" deleted successfully!')
    return redirect('config')


# Template download views
@login_required
def download_raw_materials_template(request):
    """Download raw materials template"""
    wb = ExcelTemplateGenerator.create_raw_materials_template()

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=raw_materials_template.xlsx'
    wb.save(response)
    return response


@login_required
def download_moulding_machines_template(request):
    """Download moulding machines template"""
    wb = ExcelTemplateGenerator.create_moulding_machines_template()

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=moulding_machines_template.xlsx'
    wb.save(response)
    return response


@login_required
def download_complete_quote_template(request):
    """Download complete quote template with all sheets"""
    wb = ExcelTemplateGenerator.create_complete_quote_template()

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=complete_quote_template.xlsx'
    wb.save(response)
    return response


# Upload views
@login_required
def upload_raw_materials(request, project_id, quote_id):
    """Upload raw materials from Excel"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST' and request.FILES.get('excel_file'):
        try:
            excel_file = request.FILES['excel_file']
            count, errors = ExcelParser.parse_raw_materials(excel_file, quote)

            if errors:
                for error in errors:
                    messages.error(request, error)
            else:
                quote.increment_version(request.user, f'{count} raw materials uploaded from Excel')
                messages.success(request, f'Successfully uploaded {count} raw materials!')
                return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except Exception as e:
            messages.error(request, f'Error processing file: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
    }
    return render(request, 'core/upload_raw_materials.html', context)


@login_required
def upload_moulding_machines(request, project_id, quote_id):
    """Upload moulding machines from Excel"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST' and request.FILES.get('excel_file'):
        try:
            excel_file = request.FILES['excel_file']
            count, errors = ExcelParser.parse_moulding_machines(excel_file, quote)

            if errors:
                for error in errors:
                    messages.error(request, error)
            else:
                quote.increment_version(request.user, f'{count} moulding machines uploaded from Excel')
                messages.success(request, f'Successfully uploaded {count} moulding machines!')
                return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except Exception as e:
            messages.error(request, f'Error processing file: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
    }
    return render(request, 'core/upload_moulding_machines.html', context)


@login_required
def upload_complete_quote(request, project_id, quote_id):
    """Upload complete quote from Excel with multiple sheets"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST' and request.FILES.get('excel_file'):
        try:
            excel_file = request.FILES['excel_file']
            results = ExcelParser.parse_complete_quote(excel_file, quote)

            # Show results for each section
            for section, data in results.items():
                if data['errors']:
                    for error in data['errors']:
                        messages.error(request, f"{section}: {error}")
                elif data['count'] > 0:
                    messages.success(request, f"{section}: {data['count']} items uploaded")

            quote.increment_version(request.user, 'Complete quote uploaded from Excel')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except Exception as e:
            messages.error(request, f'Error processing file: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
    }
    return render(request, 'core/upload_complete_quote.html', context)
