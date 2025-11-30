from django.urls import path
from . import views

urlpatterns = [
    path('login/', views.login_view, name='login'),
    path('ajax-login/', views.ajax_login, name='ajax_login'),
    path('forgot-password/', views.forgot_password, name='forgot_password'),
    path('verify-code/<int:user_id>/', views.verify_code, name='verify_code'),
    path('reset-password/<int:user_id>/<str:code>/', views.reset_password, name='reset_password'),
    path('logout/', views.logout_view, name='logout'),
    
]
