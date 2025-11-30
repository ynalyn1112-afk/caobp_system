from django.db import models
from django.contrib.auth import get_user_model
from django.utils import timezone
import uuid

User = get_user_model()

class OPBRequest(models.Model):
    """Operational Plan and Budget Request"""
    STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('for-approval', 'For Approval'),
        ('enhancement', 'Enhancement'),
    ]
    
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    department_head = models.ForeignKey(User, on_delete=models.CASCADE, related_name='opb_requests')
    department = models.CharField(max_length=50, choices=User.UNIT_CHOICES)
    fiscal_year = models.CharField(max_length=4, default='2026')
    unit = models.CharField(max_length=200, blank=True, null=True)
    
    # Status and Admin
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    admin_notes = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    @property
    def total_budget_amount(self):
        return sum(item.budget_amount for item in self.items.all())
    
    @property
    def item_count(self):
        return self.items.count()
    
    def __str__(self):
        return f"OPB Request - {self.unit} ({self.department}) - {self.fiscal_year}"


class OPBItem(models.Model):
    """Individual items for OPB requests"""
    request = models.ForeignKey(OPBRequest, on_delete=models.CASCADE, related_name='items')
    kra_no = models.CharField(max_length=50, blank=True, null=True)
    objective_no = models.CharField(max_length=50, blank=True, null=True)
    indicators = models.TextField(blank=True, null=True)
    annual_target = models.TextField(blank=True, null=True)
    activities = models.TextField(blank=True, null=True)
    timeframe = models.CharField(max_length=100, blank=True, null=True)
    budget_amount = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    source_of_fund = models.CharField(max_length=200, blank=True, null=True)
    responsible_units = models.CharField(max_length=200, blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return f"OPB Item - {self.kra_no} ({self.request.unit})"


class Notification(models.Model):
    """System notifications"""
    NOTIFICATION_TYPES = [
        ('info', 'Information'),
        ('success', 'Success'),
        ('warning', 'Warning'),
        ('error', 'Error'),
    ]
    
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='notifications')
    title = models.CharField(max_length=200)
    message = models.TextField()
    notification_type = models.CharField(max_length=20, choices=NOTIFICATION_TYPES, default='info')
    is_read = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return f"{self.title} - {self.user.username}"
