
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
            quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" added', 'raw_material_added')

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

    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Raw material "{material_name}" deleted', 'raw_material_deleted')

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

        # Increment version and add timeline entry
        quote.increment_version(request.user, f'Raw material section marked as complete', 'section_completed')


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
                mtc_count=int(request.POST.get('mtc_count', 0)),
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

    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Moulding machine with {cavity} cavities deleted', 'moulding_machine_deleted')

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

        # Increment version and add timeline entry
        quote.increment_version(request.user, f'Moulding machine section marked as complete', 'section_completed')

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
    assembly_types = AssemblyType.objects.filter(
        customer_group=quote.client_group,
        is_active=True
    ) if quote.client_group else []

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            assembly_type_id = request.POST.get('assembly_type_config')

            assembly = Assembly.objects.create(
                quote=quote,
                assembly_type_config_id=assembly_type_id if assembly_type_id else None,
                name=request.POST.get('name', ''),
                remarks=request.POST.get('remarks', ''),
                manual_cost=float(request.POST.get('manual_cost', 0)),
                other_cost=float(request.POST.get('other_cost', 0)),
                profit_percentage=float(request.POST.get('profit_percentage', 0)),
                rejection_percentage=float(request.POST.get('rejection_percentage', 0)),
                inspection_handling_cost=float(request.POST.get('inspection_handling_cost', 0)),
            )

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Assembly "{assembly.name}" added')

            messages.success(request, f'Assembly "{assembly.name}" added successfully!')
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
def assembly_edit(request, project_id, quote_id, assembly_id):
    """Edit assembly"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly = get_object_or_404(Assembly, id=assembly_id, quote=quote)
    assembly_types = AssemblyType.objects.filter(
        customer_group=quote.client_group,
        is_active=True
    ) if quote.client_group else []

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            assembly_type_id = request.POST.get('assembly_type_config')

            assembly.assembly_type_config_id = assembly_type_id if assembly_type_id else None
            assembly.name = request.POST.get('name', '')
            assembly.remarks = request.POST.get('remarks', '')
            assembly.manual_cost = float(request.POST.get('manual_cost', 0))
            assembly.other_cost = float(request.POST.get('other_cost', 0))
            assembly.profit_percentage = float(request.POST.get('profit_percentage', 0))
            assembly.rejection_percentage = float(request.POST.get('rejection_percentage', 0))
            assembly.inspection_handling_cost = float(request.POST.get('inspection_handling_cost', 0))
            assembly.save()

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Assembly "{assembly.name}" updated')

            messages.success(request, f'Assembly "{assembly.name}" updated successfully!')
            return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error updating assembly: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'assembly': assembly,
        'assembly_types': assembly_types,
    }
    return render(request, 'core/assembly_edit.html', context)


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

    # Increment version and add timeline entry
    quote.increment_version(request.user, f'{assembly_type} assembly deleted', 'assembly_deleted')

    messages.success(request, 'Assembly deleted successfully!')
    return redirect('quote_detail', project_id=project.id, quote_id=quote.id)


@login_required
def assembly_raw_material_add(request, project_id, quote_id, assembly_id):
    """Add raw material to assembly"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly = get_object_or_404(Assembly, id=assembly_id, quote=quote)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)

    if request.method == 'POST':
        try:
            AssemblyRawMaterial.objects.create(
                assembly=assembly,
                description=request.POST.get('description'),
                production_quantity=int(request.POST.get('production_quantity', 1)),
                production_weight=request.POST.get('production_weight', ''),
                unit=request.POST.get('unit', 'kg'),
                cost_per_unit=float(request.POST.get('cost_per_unit', 0)),
            )

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Assembly RM added to "{assembly.name}"')

            messages.success(request, 'Assembly raw material added successfully!')
            return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error adding assembly raw material: {str(e)}')

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

    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Assembly raw material "{description}" deleted from assembly #{assembly.id}', 'assembly_rm_deleted')

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

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Manufacturing cost "{mpc.process}" added to assembly #{assembly.id}', 'manufacturing_cost_added')

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

    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Manufacturing cost "{process}" deleted from assembly #{assembly.id}', 'manufacturing_cost_deleted')

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

        # Increment version and add timeline entry
        quote.increment_version(request.user, f'Assembly section marked as complete', 'section_completed')

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

    # Get packaging types for this customer group
    packaging_types = PackagingType.objects.filter(
        customer_group=quote.client_group,
        is_active=True
    ) if quote.client_group else []

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            packaging_type_id = request.POST.get('packaging_type')

            # Create packaging
            packaging = Packaging(
                quote=quote,
                packaging_type_id=packaging_type_id if packaging_type_id else None,  # Optional
                packaging_length=float(request.POST.get('packaging_length', 0)),
                packaging_breadth=float(request.POST.get('packaging_breadth', 0)),
                packaging_height=float(request.POST.get('packaging_height', 0)),
                polybag_length=float(request.POST.get('polybag_length', 0)),
                polybag_width=float(request.POST.get('polybag_width', 0)),
                lifecycle=int(request.POST.get('lifecycle', 0)),
                cost=float(request.POST.get('cost', 0)),
                maintenance_percentage=float(request.POST.get('maintenance_percentage', 0)),
                parts_per_polybag=int(request.POST.get('parts_per_polybag', 0)),
            )
            packaging.save()  # This will auto-populate dimensions from packaging_type if selected

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Packaging added')

            messages.success(request, 'Packaging added successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error adding packaging: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'packaging_types': packaging_types,
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

    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Packaging "{packaging_type}" deleted', 'packaging_deleted')

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

        # Increment version and add timeline entry
        quote.increment_version(request.user, f'Packaging section marked as complete', 'section_completed')

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

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Transport for {transport.packaging.get_packaging_type_display()} added', 'transport_added')

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

    # Increment version and add timeline entry
    quote.increment_version(request.user, f'Transport for {packaging_type} deleted', 'transport_deleted')

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
        # Increment version and add timeline entry
        quote.increment_version(request.user, f'Transport section marked as complete', 'section_completed')

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
            quote.increment_version(request.user, description, f'manual_entry', False)

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

    # Map config_type to model
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
        try:
            # Handle PackagingType separately since it has different fields
            if config_type == 'packaging-type':
                name = request.POST.get('name')
                if not name:
                    messages.error(request, 'Name is required.')
                else:
                    item = PackagingType.objects.create(
                        customer_group=customer_group,
                        name=name,
                        packaging_type=request.POST.get('packaging_type', 'custom'),
                        default_length=float(request.POST.get('default_length', 600)),
                        default_breadth=float(request.POST.get('default_breadth', 400)),
                        default_height=float(request.POST.get('default_height', 250)),
                        default_polybag_length=float(request.POST.get('default_polybag_length', 16)),
                        default_polybag_width=float(request.POST.get('default_polybag_width', 20)),
                        is_active=True
                    )
                    messages.success(request, f'{title} "{item.name}" created successfully!')
                    return redirect(f'/config/?customer_group={customer_group_id}')

            # Handle AssemblyType (with value, description, created_by)
            elif config_type == 'assembly-type':
                name = request.POST.get('name')
                value = request.POST.get('value')
                description = request.POST.get('description')

                if name and value:
                    item = AssemblyType.objects.create(
                        customer_group=customer_group,
                        name=name,
                        value=value,
                        description=description,
                        created_by=request.user
                    )
                    messages.success(request, f'{title} "{item.name}" created successfully!')
                    return redirect(f'/config/?customer_group={customer_group_id}')
                else:
                    messages.error(request, 'Name and Value are required.')

        except Exception as e:
            messages.error(request, f'Error creating {title}: {str(e)}')

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
        mtc_count = request.POST.get('mtc_count')

        if name and shift_rate and shift_rate_for_mtc and mtc_count:
            try:
                machine_type = MouldingMachineType.objects.create(
                    customer_group=customer_group,
                    name=name,
                    shift_rate=float(shift_rate),
                    shift_rate_for_mtc=float(shift_rate_for_mtc),
                    mtc_count=int(mtc_count),
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
                quote.increment_version(request.user, f'{count} raw materials uploaded from Excel', 'quote_updated')
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
                quote.increment_version(request.user, f'{count} moulding machines uploaded from Excel', 'quote_updated')
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

            quote.increment_version(request.user, 'Complete quote uploaded from Excel', 'quote_updated')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except Exception as e:
            messages.error(request, f'Error processing file: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
    }
    return render(request, 'core/upload_complete_quote.html', context)

@login_required
def raw_material_edit(request, project_id, quote_id, rm_id):
    """Edit raw material"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    raw_material = get_object_or_404(RawMaterial, id=rm_id, quote=quote)

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

            raw_material.material_type_id = material_type_id if material_type_id else None
            raw_material.material_name = request.POST.get('material_name')
            raw_material.grade = request.POST.get('grade')
            raw_material.rm_code = request.POST.get('rm_code')
            raw_material.unit_of_measurement = request.POST.get('unit_of_measurement', 'kg')
            raw_material.frozen_rate = frozen_rate
            raw_material.rm_rate = float(request.POST.get('rm_rate'))
            raw_material.part_weight = float(request.POST.get('part_weight'))
            raw_material.runner_weight = float(request.POST.get('runner_weight'))
            raw_material.process_losses = float(request.POST.get('process_losses', 0))
            raw_material.purging_loss_cost = float(request.POST.get('purging_loss_cost', 0))
            raw_material.icc_percentage = float(request.POST.get('icc_percentage', 0))
            raw_material.rejection_percentage = float(request.POST.get('rejection_percentage', 0))
            raw_material.overhead_percentage = float(request.POST.get('overhead_percentage', 0))
            raw_material.maintenance_percentage = float(request.POST.get('maintenance_percentage', 0))
            raw_material.profit_percentage = float(request.POST.get('profit_percentage', 0))

            raw_material.save()

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Raw material "{raw_material.material_name}" updated')

            messages.success(request, f'Raw material "{raw_material.material_name}" updated successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error updating raw material: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'raw_material': raw_material,
        'material_types': material_types,
        'unit_choices': RawMaterial.UNIT_CHOICES,
    }
    return render(request, 'core/raw_material_edit.html', context)


@login_required
def moulding_machine_edit(request, project_id, quote_id, mm_id):
    """Edit moulding machine"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    machine = get_object_or_404(MouldingMachineDetail, id=mm_id, quote=quote)

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

            machine.moulding_machine_type_id = moulding_machine_type_id if moulding_machine_type_id else None
            machine.cavity = int(request.POST.get('cavity', 1))
            machine.machine_tonnage = float(request.POST.get('machine_tonnage', 0))
            machine.cycle_time = float(request.POST.get('cycle_time', 0))
            machine.efficiency = float(request.POST.get('efficiency', 0))
            machine.shift_rate = float(request.POST.get('shift_rate', 0))
            machine.shift_rate_for_mtc = float(request.POST.get('shift_rate_for_mtc', 0))
            machine.mtc_count = int(request.POST.get('mtc_count', 0))
            machine.rejection_percentage = float(request.POST.get('rejection_percentage', 0))
            machine.overhead_percentage = float(request.POST.get('overhead_percentage', 0))
            machine.maintenance_percentage = float(request.POST.get('maintenance_percentage', 0))
            machine.profit_percentage = float(request.POST.get('profit_percentage', 0))

            machine.save()

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Moulding machine with {machine.cavity} cavity updated')

            messages.success(request, 'Moulding machine updated successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error updating moulding machine: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'machine': machine,
        'moulding_machine_types': moulding_machine_types,
    }
    return render(request, 'core/moulding_machine_edit.html', context)


@login_required
def assembly_raw_material_edit(request, project_id, quote_id, assembly_id, arm_id):
    """Edit assembly raw material"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly = get_object_or_404(Assembly, id=assembly_id, quote=quote)
    assembly_rm = get_object_or_404(AssemblyRawMaterial, id=arm_id, assembly=assembly)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)

    if request.method == 'POST':
        try:
            assembly_rm.description = request.POST.get('description')
            assembly_rm.production_quantity = int(request.POST.get('production_quantity', 1))
            assembly_rm.production_weight = request.POST.get('production_weight', '')
            assembly_rm.unit = request.POST.get('unit', 'kg')
            assembly_rm.cost_per_unit = float(request.POST.get('cost_per_unit', 0))
            assembly_rm.save()

            # Recalculate assembly costs
            assembly.save()

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Assembly RM "{assembly_rm.description}" updated in "{assembly.name}"')

            messages.success(request, f'Assembly raw material "{assembly_rm.description}" updated successfully!')
            return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error updating assembly raw material: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'assembly': assembly,
        'assembly_rm': assembly_rm,
        'unit_choices': AssemblyRawMaterial.UNIT_CHOICES,
    }
    return render(request, 'core/assembly_raw_material_edit.html', context)

@login_required
def manufacturing_printing_cost_edit(request, project_id, quote_id, assembly_id, cost_id):
    """Edit manufacturing/printing cost"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    assembly = get_object_or_404(Assembly, id=assembly_id, quote=quote)
    mfg_cost = get_object_or_404(ManufacturingPrintingCost, id=cost_id, assembly=assembly)

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)

    if request.method == 'POST':
        try:
            mfg_cost.process = request.POST.get('process')
            mfg_cost.mc_tonnage = float(request.POST.get('mc_tonnage', 0))
            mfg_cost.mc_rate_per_hour = float(request.POST.get('mc_rate_per_hour', 0))
            mfg_cost.cycle_time = float(request.POST.get('cycle_time', 0))
            mfg_cost.save()

            # Recalculate assembly costs
            assembly.save()

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Manufacturing cost "{mfg_cost.process}" updated in "{assembly.name}"')

            messages.success(request, f'Manufacturing/Printing cost "{mfg_cost.process}" updated successfully!')
            return redirect('assembly_detail', project_id=project.id, quote_id=quote.id, assembly_id=assembly.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error updating manufacturing cost: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'assembly': assembly,
        'mfg_cost': mfg_cost,
    }
    return render(request, 'core/manufacturing_printing_cost_edit.html', context)


@login_required
def packaging_edit(request, project_id, quote_id, packaging_id):
    """Edit packaging"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    packaging = get_object_or_404(Packaging, id=packaging_id, quote=quote)

    # Get packaging types for this customer group
    packaging_types = PackagingType.objects.filter(
        customer_group=quote.client_group,
        is_active=True
    ) if quote.client_group else []

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            packaging_type_id = request.POST.get('packaging_type')

            packaging.packaging_type_id = packaging_type_id if packaging_type_id else None
            packaging.packaging_length = float(request.POST.get('packaging_length', 0))
            packaging.packaging_breadth = float(request.POST.get('packaging_breadth', 0))
            packaging.packaging_height = float(request.POST.get('packaging_height', 0))
            packaging.polybag_length = float(request.POST.get('polybag_length', 0))
            packaging.polybag_width = float(request.POST.get('polybag_width', 0))
            packaging.lifecycle = int(request.POST.get('lifecycle', 0))
            packaging.cost = float(request.POST.get('cost', 0))
            packaging.maintenance_percentage = float(request.POST.get('maintenance_percentage', 0))
            packaging.parts_per_polybag = int(request.POST.get('parts_per_polybag', 0))
            packaging.save()

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Packaging updated')

            messages.success(request, 'Packaging updated successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error updating packaging: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'packaging': packaging,
        'packaging_types': packaging_types,
    }
    return render(request, 'core/packaging_edit.html', context)

@login_required
def download_multiple_quotes_template(request):
    """Download template for creating multiple quotes"""
    wb = ExcelTemplateGenerator.create_multiple_quotes_template()

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=multiple_quotes_template.xlsx'
    wb.save(response)
    return response


@login_required
def upload_multiple_quotes(request, project_id):
    """Upload multiple complete quotes to a project from Excel"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    customer_groups = CustomerGroup.objects.filter(is_active=True)

    if request.method == 'POST' and request.FILES.get('excel_file'):
        customer_group_id = request.POST.get('customer_group')

        if not customer_group_id:
            messages.error(request, 'Please select a customer group.')
        else:
            try:
                customer_group = get_object_or_404(CustomerGroup, id=customer_group_id)
                excel_file = request.FILES['excel_file']

                results = ExcelParser.parse_multiple_quotes_complete(
                    excel_file, project, customer_group, request.user
                )

                if results['errors']:
                    for error in results['errors']:
                        messages.error(request, error)

                if results['quotes'] > 0:
                    messages.success(request, f'Successfully created {results["quotes"]} quotes!')

                    # Show component counts
                    components = results['components']
                    if components['raw_materials'] > 0:
                        messages.info(request, f'Added {components["raw_materials"]} raw materials')
                    if components['moulding_machines'] > 0:
                        messages.info(request, f'Added {components["moulding_machines"]} moulding machines')
                    if components['assemblies'] > 0:
                        messages.info(request, f'Added {components["assemblies"]} assemblies')
                    if components['packaging'] > 0:
                        messages.info(request, f'Added {components["packaging"]} packaging entries')
                    if components['transport'] > 0:
                        messages.info(request, f'Added {components["transport"]} transport entries')

                    return redirect('project_detail', project_id=project.id)
                else:
                    messages.warning(request, 'No quotes were created. Please check your file.')

            except Exception as e:
                messages.error(request, f'Error processing file: {str(e)}')

    context = {
        'project': project,
        'customer_groups': customer_groups,
    }
    return render(request, 'core/upload_multiple_quotes.html', context)


@login_required
def transport_edit(request, project_id, quote_id, transport_id):
    """Edit transport"""
    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)
    transport = get_object_or_404(Transport, id=transport_id, quote=quote)
    packagings = quote.packagings.all()

    # Check if quote can be edited
    if not quote.can_edit_sections():
        messages.error(request, 'This quote is completed or discarded and cannot be edited. Reopen it to make changes.')
        return redirect('quote_detail', project_id=project.id, quote_id=quote.id)

    if request.method == 'POST':
        try:
            packaging_id = request.POST.get('packaging')

            transport.packaging_id = packaging_id if packaging_id else None
            transport.transport_length = float(request.POST.get('transport_length', 0))
            transport.transport_breadth = float(request.POST.get('transport_breadth', 0))
            transport.transport_height = float(request.POST.get('transport_height', 0))
            transport.trip_cost = float(request.POST.get('trip_cost', 0))
            transport.parts_per_box = int(request.POST.get('parts_per_box', 1))
            transport.save()

            # Increment version and add timeline entry
            quote.increment_version(request.user, f'Transport updated')

            messages.success(request, 'Transport updated successfully!')
            return redirect('quote_detail', project_id=project.id, quote_id=quote.id)
        except ValueError as e:
            messages.error(request, f'Invalid numeric value provided. Please check your inputs.')
        except Exception as e:
            messages.error(request, f'Error updating transport: {str(e)}')

    context = {
        'project': project,
        'quote': quote,
        'transport': transport,
        'packagings': packagings,
    }
    return render(request, 'core/transport_edit.html', context)


@login_required
def packaging_type_add(request, customer_group_id):
    """Add packaging type to customer group"""
    customer_group = get_object_or_404(CustomerGroup, id=customer_group_id, is_active=True)

    if request.method == 'POST':
        try:
            PackagingType.objects.create(
                customer_group=customer_group,
                name=request.POST.get('name'),
                packaging_type=request.POST.get('packaging_type', 'custom'),
                default_length=float(request.POST.get('default_length', 600)),
                default_breadth=float(request.POST.get('default_breadth', 400)),
                default_height=float(request.POST.get('default_height', 250)),
                default_polybag_length=float(request.POST.get('default_polybag_length', 16)),
                default_polybag_width=float(request.POST.get('default_polybag_width', 20)),
                is_active=True
            )

            messages.success(request, f'Packaging type "{request.POST.get("name")}" created successfully!')
            return redirect('config')  # or wherever you want to redirect
        except Exception as e:
            messages.error(request, f'Error creating Packaging Type: {str(e)}')

    context = {
        'customer_group': customer_group,
    }
    return render(request, 'core/packaging_type_add.html', context)


@login_required
def packaging_type_edit(request, packaging_type_id):
    """Edit packaging type"""
    packaging_type = get_object_or_404(PackagingType, id=packaging_type_id, is_active=True)

    if request.method == 'POST':
        try:
            packaging_type.name = request.POST.get('name')
            packaging_type.packaging_type = request.POST.get('packaging_type', 'custom')
            packaging_type.default_length = float(request.POST.get('default_length', 600))
            packaging_type.default_breadth = float(request.POST.get('default_breadth', 400))
            packaging_type.default_height = float(request.POST.get('default_height', 250))
            packaging_type.default_polybag_length = float(request.POST.get('default_polybag_length', 16))
            packaging_type.default_polybag_width = float(request.POST.get('default_polybag_width', 20))
            packaging_type.save()

            messages.success(request, f'Packaging type "{packaging_type.name}" updated successfully!')
            return redirect('config')
        except Exception as e:
            messages.error(request, f'Error updating Packaging Type: {str(e)}')

    context = {
        'packaging_type': packaging_type,
    }
    return render(request, 'core/packaging_type_edit.html', context)


@login_required
def packaging_type_delete(request, packaging_type_id):
    """Delete (deactivate) packaging type"""
    packaging_type = get_object_or_404(PackagingType, id=packaging_type_id)

    if request.method == 'POST':
        packaging_type.is_active = False
        packaging_type.save()

        messages.success(request, f'Packaging type "{packaging_type.name}" deactivated successfully!')
        return redirect('config')

    context = {
        'packaging_type': packaging_type,
    }
    return render(request, 'core/packaging_type_delete.html', context)


@login_required
def export_quote(request, project_id, quote_id):
    """Export quote to Excel with all calculated fields"""
    from core.excel_utils import ExcelExporter

    project = get_object_or_404(Project, id=project_id, is_active=True)
    quote = get_object_or_404(Quote, id=quote_id, project=project)

    # Generate Excel file
    wb = ExcelExporter.export_quote(quote)

    # Create response
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    # Get version string for filename
    version_str = f"{quote.version_major}.{quote.version_minor}" if hasattr(quote, 'version_major') else ''
    filename = f'Quote_{quote.name}_{version_str}.xlsx'.replace(' ', '_')
    response['Content-Disposition'] = f'attachment; filename={filename}'

    wb.save(response)
    return response

@login_required
def export_project(request, project_id):
    """Export entire project with all quotes to Excel"""
    from core.excel_utils import ExcelExporter

    project = get_object_or_404(Project, id=project_id, is_active=True)

    # Generate Excel file
    wb = ExcelExporter.export_project(project)

    # Create response
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    filename = f'Project_{project.name}.xlsx'.replace(' ', '_')
    response['Content-Disposition'] = f'attachment; filename={filename}'

    wb.save(response)
    return response
