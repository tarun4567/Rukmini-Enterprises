from django.urls import path
from . import views

urlpatterns = [
    path('', views.dashboard_view, name='dashboard'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    
    # Stock (Add) & Stock 1 (List) routes for Admin
    path('stock/', views.stock_add_view, name='stock_add'),
    path('stock1/', views.stock1_view, name='stock1'),
    path('stock/edit/<int:pk>/', views.stock_edit_view, name='stock_edit'),
    path('stock/delete/<int:pk>/', views.stock_delete_view, name='stock_delete'),

    path('stocks/', views.stocks_view, name='stocks'),
    path('billing/', views.billing_view, name='billing'),
    path('billing/records/', views.billing_records_view, name='billing_records'),
    path('billing/clear-due/<int:pk>/', views.clear_due_view, name='clear_due'),
]
