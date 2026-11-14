from django.contrib import admin
from .models import Student, Attendance, Curator, Headman

@admin.register(Student)
class StudentAdmin(admin.ModelAdmin):
    list_display = ('full_name', 'group_name')

@admin.register(Attendance)
class AttendanceAdmin(admin.ModelAdmin):
    list_display = ('student', 'date', 'status', 'color')
    list_filter = ('status', 'date')
    search_fields = ('student__full_name',)

@admin.register(Curator)
class CuratorAdmin(admin.ModelAdmin):
    list_display = ('full_name', 'group_name')

@admin.register(Headman)
class HeadmanAdmin(admin.ModelAdmin):
    list_display = ('full_name', 'group_name')
