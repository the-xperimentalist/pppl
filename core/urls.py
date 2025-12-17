from django.urls import path
from . import views


urlpatterns = [
    path('', views.home, name='home'),
    path('projects/', views.projects, name='projects'),
    path('projects/new/', views.project_create, name='project_create'),
    path('projects/<int:project_id>/', views.project_detail, name='project_detail'),
    path('projects/<int:project_id>/quotes/new/', views.quote_create, name='quote_create'),

    # Quote detail and sections
    path('projects/<int:project_id>/quotes/<int:quote_id>/', views.quote_detail, name='quote_detail'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/definition/edit/', views.quote_definition_edit, name='quote_definition_edit'),

    # Raw Materials
    path('projects/<int:project_id>/quotes/<int:quote_id>/raw-materials/add/', views.raw_material_add, name='raw_material_add'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/raw-materials/<int:rm_id>/delete/', views.raw_material_delete, name='raw_material_delete'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/raw-materials/complete/', views.raw_material_complete, name='raw_material_complete'),

    # Moulding Machines
    path('projects/<int:project_id>/quotes/<int:quote_id>/moulding-machines/add/', views.moulding_machine_add, name='moulding_machine_add'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/moulding-machines/<int:mm_id>/delete/', views.moulding_machine_delete, name='moulding_machine_delete'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/moulding-machines/complete/', views.moulding_machine_complete, name='moulding_machine_complete'),
    # Config - Customer Groups
    path('config/customer-groups/create/', views.customer_group_create, name='customer_group_create'),
    path('config/customer-groups/<int:customer_group_id>/delete/', views.customer_group_delete, name='customer_group_delete'),

    # Config
    path('config/', views.config, name='config'),
    path('config/<str:config_type>/new/', views.config_create, name='config_create'),
    path('config/<str:config_type>/<int:item_id>/delete/', views.config_delete, name='config_delete'),
    # Assembly
    path('projects/<int:project_id>/quotes/<int:quote_id>/assemblies/add/', views.assembly_add, name='assembly_add'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/assemblies/<int:assembly_id>/', views.assembly_detail, name='assembly_detail'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/assemblies/<int:assembly_id>/delete/', views.assembly_delete, name='assembly_delete'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/assemblies/<int:assembly_id>/raw-materials/add/', views.assembly_raw_material_add, name='assembly_raw_material_add'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/assemblies/<int:assembly_id>/raw-materials/<int:arm_id>/delete/', views.assembly_raw_material_delete, name='assembly_raw_material_delete'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/assemblies/<int:assembly_id>/manufacturing-costs/add/', views.manufacturing_printing_cost_add, name='manufacturing_printing_cost_add'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/assemblies/<int:assembly_id>/manufacturing-costs/<int:mpc_id>/delete/', views.manufacturing_printing_cost_delete, name='manufacturing_printing_cost_delete'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/assemblies/complete/', views.assembly_complete, name='assembly_complete'),
    # Packaging
    path('projects/<int:project_id>/quotes/<int:quote_id>/packaging/add/', views.packaging_add, name='packaging_add'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/packaging/<int:packaging_id>/delete/', views.packaging_delete, name='packaging_delete'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/packaging/complete/', views.packaging_complete, name='packaging_complete'),
    # Transport
    path('projects/<int:project_id>/quotes/<int:quote_id>/transport/add/', views.transport_add, name='transport_add'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/transport/<int:transport_id>/delete/', views.transport_delete, name='transport_delete'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/transport/complete/', views.transport_complete, name='transport_complete'),
    # Timeline
    path('projects/<int:project_id>/quotes/<int:quote_id>/timeline/add/', views.timeline_add_manual, name='timeline_add_manual'),
    # Quote summary
    path('projects/<int:project_id>/quotes/<int:quote_id>/summary/', views.quote_summary, name='quote_summary'),
    # Admin Dashboard (superuser only)
    path('admin-dashboard/', views.admin_dashboard, name='admin_dashboard'),
    path('admin-dashboard/users/create/', views.admin_user_create, name='admin_user_create'),
    path('admin-dashboard/users/<int:user_id>/edit/', views.admin_user_edit, name='admin_user_edit'),
    path('admin-dashboard/users/<int:user_id>/delete/', views.admin_user_delete, name='admin_user_delete'),
    path('admin-dashboard/users/<int:user_id>/toggle-active/', views.admin_user_toggle_active, name='admin_user_toggle_active'),
    # Quote status management
    path('projects/<int:project_id>/quotes/<int:quote_id>/mark-completed/', views.quote_mark_completed, name='quote_mark_completed'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/reopen/', views.quote_reopen, name='quote_reopen'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/discard/', views.quote_discard, name='quote_discard'),
    # Material Types
    path('config/customer-groups/<int:customer_group_id>/material-types/create/', views.material_type_create, name='material_type_create'),
    path('config/material-types/<int:material_type_id>/delete/', views.material_type_delete, name='material_type_delete'),

    # Moulding Machine Types
    path('config/customer-groups/<int:customer_group_id>/moulding-machine-types/create/', views.moulding_machine_type_create, name='moulding_machine_type_create'),
    path('config/moulding-machine-types/<int:machine_type_id>/delete/', views.moulding_machine_type_delete, name='moulding_machine_type_delete'),
    # Excel template downloads
    path('templates/raw-materials/', views.download_raw_materials_template, name='download_raw_materials_template'),
    path('templates/moulding-machines/', views.download_moulding_machines_template, name='download_moulding_machines_template'),
    path('templates/complete-quote/', views.download_complete_quote_template, name='download_complete_quote_template'),

    # Excel uploads
    path('projects/<int:project_id>/quotes/<int:quote_id>/upload/raw-materials/', views.upload_raw_materials, name='upload_raw_materials'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/upload/moulding-machines/', views.upload_moulding_machines, name='upload_moulding_machines'),
    path('projects/<int:project_id>/quotes/<int:quote_id>/upload/complete/', views.upload_complete_quote, name='upload_complete_quote'),
    # Raw Material edit
    path('projects/<int:project_id>/quotes/<int:quote_id>/raw-materials/<int:rm_id>/edit/', views.raw_material_edit, name='raw_material_edit'),

    # Moulding Machine edit
    path('projects/<int:project_id>/quotes/<int:quote_id>/moulding-machines/<int:mm_id>/edit/', views.moulding_machine_edit, name='moulding_machine_edit'),
]