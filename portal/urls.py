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
    path('billing/dashboard/download/excel/', views.user_dashboard_excel, name='user_dashboard_excel'),
    path('billing/dashboard/download/pdf/',   views.user_dashboard_pdf,   name='user_dashboard_pdf'),
    path('billing/clear-due/<int:pk>/', views.clear_due_view, name='clear_due'),
    path('billing/report/', views.revenue_report_view, name='revenue_report'),
    path('billing/expenses-ledger/', views.admin_expenses_view, name='admin_expenses'),

    # Download endpoints
    path('billing/report/download/excel/', views.revenue_report_excel, name='revenue_report_excel'),
    path('billing/report/download/pdf/',   views.revenue_report_pdf,   name='revenue_report_pdf'),
    path('stock/report/download/excel/',   views.stock_report_excel,   name='stock_report_excel'),
    path('stock/report/download/pdf/',     views.stock_report_pdf,     name='stock_report_pdf'),

    # Stock History download (per product)
    path('stock/<int:pk>/history/excel/', views.stock_history_excel, name='stock_history_excel'),
    path('stock/<int:pk>/history/pdf/',   views.stock_history_pdf,   name='stock_history_pdf'),

    # All Stock History download (all products)
    path('stock/history/all/excel/', views.stock_all_history_excel, name='stock_all_history_excel'),
    path('stock/history/all/pdf/',   views.stock_all_history_pdf,   name='stock_all_history_pdf'),
]
