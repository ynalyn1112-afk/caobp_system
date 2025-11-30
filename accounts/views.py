from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
from django.utils import timezone
from django.core.mail import send_mail
from django.conf import settings
from django.contrib.auth.hashers import make_password
from .models import User, PasswordResetCode
import json


def landing_page(request):
    return render(request, 'landing.html')


def login_view(request):
    if request.user.is_authenticated:
        if request.user.role == 'admin':
            return redirect('admin_dashboard')
        else:
            return redirect('head_dashboard')
    
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')
        
        user = authenticate(request, username=username, password=password)
        
        if user is not None:
            if user.is_active:
                login(request, user)
                if user.role == 'admin':
                    return redirect('budget:admin_dashboard')
                else:
                    return redirect('budget:head_dashboard')
            else:
                messages.error(request, 'Account is deactivated. Please contact administrator.')
                return redirect('login')
        else:
            messages.error(request, 'Invalid credentials. Please try again.')
            return redirect('login')
    
    return render(request, 'login.html')


@csrf_exempt
@require_http_methods(["POST"])
def ajax_login(request):
    """AJAX login"""
    try:
        # Handle both JSON and FormData
        if request.content_type == 'application/json':
            data = json.loads(request.body)
            username = data.get('username')
            password = data.get('password')
        else:
            username = request.POST.get('username')
            password = request.POST.get('password')
        
        user = authenticate(request, username=username, password=password)
        
        if user is not None and user.is_active:
            login(request, user)
            return JsonResponse({
                'success': True,
                'redirect_url': '/budget/admin-dashboard/' if user.role == 'admin' else '/budget/head-dashboard/'
            })
        else:
            return JsonResponse({
                'success': False,
                'message': 'Invalid credentials or account deactivated',
                'redirect_url': '/accounts/login/'
            })
    except Exception as e:
        print(f"Login error: {e}")
        return JsonResponse({
            'success': False,
            'message': 'An error occurred. Please try again.'
        })


def forgot_password(request):
    if request.method == 'POST':
        email = request.POST.get('email')
        try:
            user = User.objects.get(email=email)
            
            # Check if user is active
            if not user.is_active:
                messages.error(request, 'This account has been deactivated. Please contact the administrator.')
                return render(request, 'forgot_password.html')
            
            # Create or update reset code
            reset_code, created = PasswordResetCode.objects.get_or_create(
                user=user,
                defaults={'expires_at': timezone.now() + timezone.timedelta(minutes=10)}
            )
            if not created:
                reset_code.expires_at = timezone.now() + timezone.timedelta(minutes=10)
                reset_code.is_used = False
                reset_code.save()
            
            # Send email with verification code
            try:
                from django.core.mail import send_mail
                
                subject = 'CAOBP System - Password Reset Verification Code'
                message = f"""
Hello {user.get_full_name() or user.username},

You have requested to reset your password for the CAOBP System.

Your verification code is: {reset_code.code}

This code will expire in 10 minutes.

If you didn't request this reset, please ignore this email.

Best regards,
CAOBP System Team
                """
                
                send_mail(
                    subject,
                    message,
                    settings.DEFAULT_FROM_EMAIL,
                    [user.email],
                    fail_silently=False,
                )
                
                messages.success(request, f'Verification code sent to {email}. Please check your email.')
                return redirect('verify_code', user_id=user.id)
                
            except Exception as e:
                # If email fail
                messages.warning(request, f'Email sending failed')
                return redirect('verify_code', user_id=user.id)
            
        except User.DoesNotExist:
            messages.error(request, 'No account found with this email address.')
            return redirect('forgot_password')
    
    return render(request, 'forgot_password.html')


def verify_code(request, user_id):
    try:
        user = User.objects.get(id=user_id)
    except User.DoesNotExist:
        messages.error(request, 'Invalid user.')
        return redirect('forgot_password')
    
    if request.method == 'POST':
        code = request.POST.get('code')
        try:
            reset_code = PasswordResetCode.objects.get(
                user=user,
                code=code,
                is_used=False
            )
            if reset_code.is_valid():
                return redirect('reset_password', user_id=user.id, code=code)
            else:
                messages.error(request, 'Code has expired. Please request a new one.')
        except PasswordResetCode.DoesNotExist:
            messages.error(request, 'Invalid verification code.')
    
    return render(request, 'verify_code.html', {'user': user})


def reset_password(request, user_id, code):
    try:
        user = User.objects.get(id=user_id)
        reset_code = PasswordResetCode.objects.get(
            user=user,
            code=code,
            is_used=False
        )
        if not reset_code.is_valid():
            messages.error(request, 'Code has expired. Please request a new one.')
            return redirect('forgot_password')
    except (User.DoesNotExist, PasswordResetCode.DoesNotExist):
        messages.error(request, 'Invalid reset link.')
        return redirect('forgot_password')
    
    if request.method == 'POST':
        new_password = request.POST.get('new_password')
        confirm_password = request.POST.get('confirm_password')
        
        if new_password != confirm_password:
            messages.error(request, 'Passwords do not match.')
        elif len(new_password) < 8:
            messages.error(request, 'Password must be at least 8 characters long.')
        else:
            user.set_password(new_password)
            user.save()
            reset_code.is_used = True
            reset_code.save()
            
            # Auto-login after password reset
            login(request, user)
            messages.success(request, 'Password reset successfully!')
            
            if user.role == 'admin':
                return redirect('admin_dashboard')
            else:
                return redirect('head_dashboard')
    
    return render(request, 'reset_password.html', {'user': user, 'code': code})


def logout_view(request):
    """Logout"""
    logout(request)
    messages.success(request, 'You have been logged out successfully.')
    return redirect('landing_page')