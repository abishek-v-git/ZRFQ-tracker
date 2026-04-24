from django.urls import path
from . import views

urlpatterns = [
    path('', views.rfq_list, name='rfq_list'),
    path('add/', views.rfq_add, name='rfq_add'),
    path('delete/<int:pk>/', views.rfq_delete, name='rfq_delete'),
    path('edit/<int:pk>/', views.rfq_edit, name='rfq_edit'),
    path('export/', views.rfq_export, name='rfq_export'),
    path('bulk-upload/', views.rfq_bulk_upload, name='rfq_bulk_upload'),
    path('clear-all/', views.rfq_clear_all, name='rfq_clear_all'),
    path('patch/<int:pk>/', views.rfq_patch, name='rfq_patch'),
    path('edit-json/<int:pk>/', views.rfq_edit_json, name='rfq_edit_json'),
    path('data/', views.rfq_data, name='rfq_data'),
    # Supplier Info
    path('suppliers/', views.supplier_list, name='supplier_list'),
    path('suppliers/data/', views.supplier_data, name='supplier_data'),
    path('suppliers/save/', views.supplier_save, name='supplier_save'),
    path('suppliers/delete/<int:pk>/', views.supplier_delete, name='supplier_delete'),
    path('suppliers/upload-template/', views.supplier_template_upload, name='supplier_template_upload'),
]
