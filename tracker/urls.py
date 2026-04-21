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
]
