from django.contrib import admin
from django.urls import path
from django.contrib.auth import views as auth_views
from django.views.generic import RedirectView
from attendance import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('login/', auth_views.LoginView.as_view(), name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='login'), name='logout'),
    
    # Наш дашборд
    path('dashboard/', views.dashboard, name='dashboard'),
    path('', RedirectView.as_view(url='login/')),
    path('save_journal_marks/', views.save_journal_marks, name='save_journal_marks'),
    path('export_excel/', views.export_excel, name='export_excel'),
]
