from django.db import models
from django.contrib.auth.models import User
from django.db.models.signals import post_save
from django.dispatch import receiver

class Student(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='student_profile')
    full_name = models.CharField(max_length=255, verbose_name="ФИО")
    group_name = models.CharField(max_length=50, verbose_name="Группа")

    def __str__(self):
        return self.full_name

class Curator(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='curator_profile')
    full_name = models.CharField(max_length=255, verbose_name="ФИО куратора")
    group_name = models.CharField(max_length=50, verbose_name="Группа")

    def __str__(self):
        return f"Куратор: {self.full_name}"

class Headman(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='headman_profile')
    full_name = models.CharField(max_length=255, verbose_name="ФИО старосты")
    group_name = models.CharField(max_length=50, verbose_name="Группа")

    def __str__(self):
        return f"Староста: {self.full_name}"

class Attendance(models.Model):
    STATUS_CHOICES = (
        ('present', 'Был'),
        ('valid_absent', 'Уважительная (2)'),
        ('invalid_absent', 'Неуважительная (4)'),
        ('sick', 'По болезни (6)'),
        ('other', 'Другое (8)'),
    )
    
    COLOR_CHOICES = (
        ('none', 'Без цвета'),
        ('red', 'Красный'),
        ('blue', 'Синий'),
        ('green', 'Зеленый'),
        ('yellow', 'Желтый'),
    )
    
    # Теперь мы храним просто дату, без привязки к занятию
    date = models.DateField(verbose_name="Дата")
    student = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='attendances')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='present', verbose_name="Статус")
    color = models.CharField(max_length=20, choices=COLOR_CHOICES, default='none', verbose_name="Цвет")

    class Meta:
        unique_together = ('date', 'student') # Один студент - одна отметка в день

    def __str__(self):
        return f"{self.student.full_name} - {self.date} - {self.get_status_display()}"

@receiver(post_save, sender=User)
def create_student_profile(sender, instance, created, **kwargs):
    if created and not instance.is_superuser:
        if not hasattr(instance, 'curator_profile') and not hasattr(instance, 'headman_profile'):
            Student.objects.create(user=instance, full_name=instance.username)