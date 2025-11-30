from django.urls import path
from . import views

urlpatterns = [
    #Admin Dashboard
    path('admin-dashboard/', views.admin_dashboard, name='admin_dashboard'),
    path('admin-users/', views.admin_users, name='admin_users'),
    path('admin-opb/', views.admin_opb_requests, name='admin_opb'),
    path('admin-opb/<uuid:request_id>/', views.admin_opb_view_details, name='admin_opb_details'),
    path('admin-reports/', views.admin_reports, name='admin_reports'),
    path('admin-settings/', views.admin_settings, name='admin_settings'),
    
    #Department Head
    path('head-dashboard/', views.head_dashboard, name='head_dashboard'),
    path('head-opb/', views.head_opb_requests, name='head_opb'),
    path('head-opb/edit/<uuid:request_id>/', views.head_opb_edit, name='head_opb_edit'),
    path('head-opb/view/<uuid:request_id>/', views.head_opb_view, name='head_opb_view'),
    path('head-notifications/', views.head_notifications, name='head_notifications'),
    
    # AJAX endpoints
    path('ajax/add-user/', views.ajax_add_user, name='ajax_add_user'),
    path('ajax/get-user/<int:user_id>/', views.ajax_get_user, name='ajax_get_user'),
    path('ajax/edit-user/', views.ajax_edit_user, name='ajax_edit_user'),
    path('ajax/delete-user/', views.ajax_delete_user, name='ajax_delete_user'),
    path('ajax/toggle-user-status/', views.ajax_toggle_user_status, name='ajax_toggle_user_status'),
    path('ajax/approve-request/', views.ajax_approve_request, name='ajax_approve_request'),
    path('ajax/reject-request/', views.ajax_reject_request, name='ajax_reject_request'),
    path('ajax/submit-opb-request/', views.ajax_submit_opb_request, name='ajax_submit_opb_request'),
    path('ajax/delete-opb-request/', views.ajax_delete_opb_request, name='ajax_delete_opb_request'),
    path('ajax/mark-notification-read/', views.ajax_mark_notification_read, name='ajax_mark_notification_read'),
    path('ajax/delete-notification/', views.ajax_delete_notification, name='ajax_delete_notification'),
    path('ajax/mark-all-notifications-read/', views.ajax_mark_all_notifications_read, name='ajax_mark_all_notifications_read'),
    path('ajax/delete-all-notifications/', views.ajax_delete_all_notifications, name='ajax_delete_all_notifications'),
    
    # Backup management endpoints
    path('ajax/create-backup/', views.ajax_create_backup, name='ajax_create_backup'),
    path('ajax/download-backup/', views.ajax_download_backup, name='ajax_download_backup'),
    path('ajax/delete-backup/', views.ajax_delete_backup, name='ajax_delete_backup'),
    path('ajax/restore-backup/', views.ajax_restore_backup, name='ajax_restore_backup'),
]
