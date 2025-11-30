from django.contrib.auth.models import AbstractUser
from django.db import models
from django.utils import timezone
import random
import string


class User(AbstractUser):
    UNIT_CHOICES = [
        ('CAS', 'College of Arts and Sciences'),
        ('COLLEGE_IND_TECH', 'College of Industrial Technology'),
        ('COE', 'College of Education'),
        ('COA', 'College of Agriculture'),
        ('PRODUCTION_COMMERCIALIZATION', 'Production & Commercialization'),
        ('PRODUCTION_BAO', 'Production & Business Affairs Office'),
        ('PPSDO', 'Physical Plant & Site Development Office'),
        ('VETERINARY_SERVICES', 'Veterinary Services'),
        ('GUIDANCE', 'Guidance'),
        ('RECORDS_ARCHIVES', 'Records and Archives'),
        ('ALUMNI_AFFAIRS', 'Alumni Affairs'),
        ('BOARD_SECRETARY', 'Board Secretary'),
        ('MEDICAL_SERVICES', 'Medical Services Unit'),
        ('SUPERVISING_ADMIN', 'Supervising Administrative Office'),
        ('REGISTRAR', 'Registrar'),
        ('CAWAYAN_CAMPUS', 'Cawayan Campus'),
        ('BUDGET_UNIT', 'Budget Unit'),
        ('ACCOUNTING_UNIT', 'Accounting Unit'),
        ('CASH_UNIT', 'Cash Unit'),
        ('SUPPLY_UNIT', 'Supply Unit'),
        ('RECORDS_UNIT', 'Records Unit'),
        ('HRMO_OFFICE', 'HRMO Office'),
        ('PROCUREMENT_UNIT', 'Procurement Unit/BAC'),
        ('SECURITY_UNIT', 'Security Unit'),
        ('MOTORPOOL_UNIT', 'Motorpool Unit'),
        ('LIBRARY_SERVICES', 'Library Services'),
        ('NSTP_ROTC', 'NSTP/ROTC'),
        ('INTERNATIONAL_RELATIONS', 'International Relations Office'),
        ('SPORTS_CULTURAL', 'Sports and Cultural Development'),
        ('LEGAL_SERVICES', 'Legal Services Office'),
        ('PERSONNEL_SCHOLARSHIP', 'Personnel Scholarship Office'),
        ('QUALITY_ASSURANCE', 'Quality Assurance Office'),
        ('OFFICE_VPAA', 'Office of the VPAA'),
        ('OFFICE_VPAF', 'Office of the VPAF'),
        ('OFFICE_VPREICWKM', 'Office of the VPREICWKM'),
        ('OFFICE_BSB', 'Office of the Board Secretary/BOT'),
        ('PLANNING', 'Planning'),
        ('QUALITY_ASSURANCE', 'Quality Assurance'),
        ('ISO', 'ISO'),
        ('MIS', 'MIS'),
        ('SCUAA', 'SCUAA'),
        ('CSC', 'CSC'),
        ('CRCYC', 'CRCYC'),
        ('GAD', 'GAD'),
        ('PIO', 'PIO'),
    ]
    
    ROLE_CHOICES = [
        ('admin', 'Administrator'),
        ('unit_head', 'Unit Head'),
    ]
    
    email = models.EmailField(unique=True)
    department = models.CharField(max_length=50, choices=UNIT_CHOICES, blank=True, null=True, unique=True)
    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default='unit_head')
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    def __str__(self):
        return f"{self.get_full_name()} ({self.username})"


class PasswordResetCode(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    code = models.CharField(max_length=6)
    created_at = models.DateTimeField(auto_now_add=True)
    expires_at = models.DateTimeField()
    is_used = models.BooleanField(default=False)
    
    def save(self, *args, **kwargs):
        if not self.code:
            self.code = ''.join(random.choices(string.digits, k=6))
        if not self.expires_at:
            self.expires_at = timezone.now() + timezone.timedelta(minutes=10)
        super().save(*args, **kwargs)
    
    def is_valid(self):
        return not self.is_used and timezone.now() < self.expires_at
    
    def __str__(self):
        return f"Reset code for {self.user.username} - {self.code}"