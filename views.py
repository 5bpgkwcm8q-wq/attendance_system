from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib.auth.models import User
from django.contrib import messages
from .models import Student, Attendance, Curator, Headman
from django.utils import timezone
from django.http import JsonResponse, HttpResponse
import calendar
import json
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

MONTHS = {
    1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель',
    5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
    9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
}

@login_required
def dashboard(request):
    now = timezone.now()
    current_month = int(request.GET.get('month', now.month))
    current_year = int(request.GET.get('year', now.year))

    is_superuser = request.user.is_superuser
    is_curator = hasattr(request.user, 'curator_profile')
    is_headman = hasattr(request.user, 'headman_profile')
    is_student = hasattr(request.user, 'student_profile')

    # --- 1. ЗАВУЧ ---
    if is_superuser:
        if request.method == 'POST' and 'delete_student' in request.POST:
            student_id = request.POST.get('student_id')
            try:
                Student.objects.get(id=student_id).user.delete()
                messages.success(request, 'Студент удален!')
            except Student.DoesNotExist:
                pass
            return redirect(f'/dashboard/?month={current_month}&year={current_year}')

        if request.method == 'POST' and 'set_headman' in request.POST:
            group_name = request.POST.get('group_name')
            student_id = request.POST.get('student_id')
            
            old_headman = Headman.objects.filter(group_name=group_name).first()
            if old_headman:
                old_headman.delete()

            if student_id:
                try:
                    student = Student.objects.get(id=student_id)
                    headman, _ = Headman.objects.get_or_create(user=student.user)
                    headman.full_name = student.full_name
                    headman.group_name = group_name
                    headman.save()
                    messages.success(request, f'Староста группы {group_name} назначен!')
                except Student.DoesNotExist:
                    pass
            else:
                messages.success(request, f'Староста группы {group_name} снят!')
            
            return redirect(f'/dashboard/?month={current_month}&year={current_year}')

        if request.method == 'POST' and 'create_student' in request.POST:
            username = request.POST.get('username')
            password = request.POST.get('password')
            full_name = request.POST.get('full_name')
            group_name = request.POST.get('group_name')
            if User.objects.filter(username=username).exists():
                messages.error(request, 'Такой логин уже существует!')
            else:
                user = User.objects.create_user(username=username, password=password)
                student, _ = Student.objects.get_or_create(user=user)
                student.full_name = full_name
                student.group_name = group_name
                student.save()
                messages.success(request, 'Студент успешно создан!')
                return redirect(f'/dashboard/?month={current_month}&year={current_year}')

        if request.method == 'POST' and 'create_curator' in request.POST:
            username = request.POST.get('username')
            password = request.POST.get('password')
            full_name = request.POST.get('full_name')
            group_name = request.POST.get('group_name')
            if User.objects.filter(username=username).exists():
                messages.error(request, 'Такой логин уже существует!')
            else:
                user = User.objects.create_user(username=username, password=password)
                curator, _ = Curator.objects.get_or_create(user=user)
                curator.full_name = full_name
                curator.group_name = group_name
                curator.save()
                Student.objects.filter(user=user).delete()
                messages.success(request, 'Куратор успешно создан!')
                return redirect(f'/dashboard/?month={current_month}&year={current_year}')

        num_days = calendar.monthrange(current_year, current_month)[1]
        days_list = [{'num': d, 'is_sunday': date(current_year, current_month, d).weekday() == 6} for d in range(1, num_days + 1)]

        marks = {}
        all_atts = Attendance.objects.filter(date__month=current_month, date__year=current_year)
        for att in all_atts:
            day = att.date.day
            if att.status == 'invalid_absent': status_code = '4'
            elif att.status == 'valid_absent': status_code = '2'
            elif att.status == 'sick': status_code = '6'
            elif att.status == 'other': status_code = '8'
            elif att.status == 'present': status_code = '1'
            else: status_code = '0'
            marks[(att.student_id, day)] = {'status': status_code, 'color': att.color}

        groups = Student.objects.exclude(group_name__isnull=True).exclude(group_name='').values_list('group_name', flat=True).distinct().order_by('group_name')
        
        groups_data = []
        for group_name in groups:
            group_students = Student.objects.filter(group_name=group_name).order_by('full_name')
            if not group_students.exists():
                continue
                
            group_table_data = []
            for student in group_students:
                is_this_headman = hasattr(student.user, 'headman_profile')
                row_cells = []
                total_missed = 0
                valid_missed = 0
                invalid_missed = 0
                
                for day in days_list:
                    cell_data = marks.get((student.id, day['num']), {'status': '0', 'color': 'none'})
                    row_cells.append({'day': day['num'], 'is_sunday': day['is_sunday'], 'status': cell_data['status'], 'color': cell_data['color']})
                    
                    color = cell_data['color']
                    status = cell_data['status']
                    if status in ['2', '4', '6', '8']:
                        hours = int(status)
                        if color == 'red':
                            invalid_missed += hours
                            total_missed += hours
                        elif color in ['blue', 'green', 'yellow']:
                            valid_missed += hours
                            total_missed += hours
                
                group_table_data.append({
                    'student': student, 'cells': row_cells, 'is_headman': is_this_headman,
                    'total_missed': total_missed, 'valid_missed': valid_missed, 'invalid_missed': invalid_missed
                })
            
            groups_data.append({
                'name': group_name, 'students': group_table_data, 'count': group_students.count(),
                'current_headman': Headman.objects.filter(group_name=group_name).first(),
                'group_curator': Curator.objects.filter(group_name=group_name).first()
            })

        return render(request, 'dashboard_teacher.html', {
            'groups_data': groups_data, 'months': MONTHS, 'current_month': current_month, 
            'current_year': current_year, 'years': range(2024, 2030), 'days': days_list
        })

    # --- 2. СТАРОСТА ---
    elif is_headman:
        user_group = request.user.headman_profile.group_name
        role_title = f"Староста группы {user_group}"
        clean_group = user_group.strip() if user_group else ""

        students = Student.objects.filter(group_name__iexact=clean_group).order_by('full_name')
        if not students.exists():
            students = Student.objects.filter(group_name__icontains=clean_group).order_by('full_name')

        num_days = calendar.monthrange(current_year, current_month)[1]
        days_list = [{'num': d, 'is_sunday': date(current_year, current_month, d).weekday() == 6} for d in range(1, num_days + 1)]

        marks = {}
        all_atts = Attendance.objects.filter(date__month=current_month, date__year=current_year)
        for att in all_atts:
            day = att.date.day
            if att.status == 'invalid_absent': status_code = '4'
            elif att.status == 'valid_absent': status_code = '2'
            elif att.status == 'sick': status_code = '6'
            elif att.status == 'other': status_code = '8'
            elif att.status == 'present': status_code = '1'
            else: status_code = '0'
            marks[(att.student_id, day)] = {'status': status_code, 'color': att.color}

        group_table_data = []
        for student in students:
            is_this_headman = hasattr(student.user, 'headman_profile')
            row_cells = []
            total_missed = 0
            valid_missed = 0
            invalid_missed = 0
            
            for day in days_list:
                cell_data = marks.get((student.id, day['num']), {'status': '0', 'color': 'none'})
                row_cells.append({'day': day['num'], 'is_sunday': day['is_sunday'], 'status': cell_data['status'], 'color': cell_data['color']})
                
                color = cell_data['color']
                status = cell_data['status']
                if status in ['2', '4', '6', '8']:
                    hours = int(status)
                    if color == 'red':
                        invalid_missed += hours
                        total_missed += hours
                    elif color in ['blue', 'green', 'yellow']:
                        valid_missed += hours
                        total_missed += hours
                    
            group_table_data.append({
                'student': student, 'cells': row_cells, 'is_headman': is_this_headman,
                'total_missed': total_missed, 'valid_missed': valid_missed, 'invalid_missed': invalid_missed
            })

        personal_student = Student.objects.filter(user=request.user).first()
        personal_stats = None
        personal_cells = []
        if personal_student:
            p_total, p_valid, p_invalid = 0, 0, 0
            for day in days_list:
                cell_data = marks.get((personal_student.id, day['num']), {'status': '0', 'color': 'none'})
                personal_cells.append({'day': day['num'], 'is_sunday': day['is_sunday'], 'status': cell_data['status'], 'color': cell_data['color']})
                
                color = cell_data['color']
                status = cell_data['status']
                if status in ['2', '4', '6', '8']:
                    hours = int(status)
                    if color == 'red':
                        p_invalid += hours
                        p_total += hours
                    elif color in ['blue', 'green', 'yellow']:
                        p_valid += hours
                        p_total += hours
            
            personal_stats = {'name': personal_student.full_name, 'total': p_total, 'valid': p_valid, 'invalid': p_invalid}

        return render(request, 'dashboard_headman.html', {
            'group_name': clean_group, 'role_title': role_title, 'students': group_table_data,
            'group_curator': Curator.objects.filter(group_name=clean_group).first(),
            'current_headman': Headman.objects.filter(group_name=clean_group).first(),
            'personal_stats': personal_stats, 'personal_cells': personal_cells,
            'months': MONTHS, 'current_month': current_month, 'current_year': current_year,
            'years': range(2024, 2030), 'days': days_list
        })

    # --- 3. КУРАТОР ---
    elif is_curator:
        user_group = request.user.curator_profile.group_name
        role_title = f"Куратор группы {user_group}"
        clean_group = user_group.strip() if user_group else ""

        if request.method == 'POST' and 'set_headman_curator' in request.POST:
            student_id = request.POST.get('student_id')
            old_headman = Headman.objects.filter(group_name=clean_group).first()
            if old_headman:
                old_headman.delete()

            if student_id:
                try:
                    student = Student.objects.get(id=student_id)
                    headman, _ = Headman.objects.get_or_create(user=student.user)
                    headman.full_name = student.full_name
                    headman.group_name = clean_group
                    headman.save()
                    messages.success(request, f'Староста группы {clean_group} назначен!')
                except Student.DoesNotExist:
                    pass
            else:
                messages.success(request, f'Староста группы {clean_group} снят!')
            
            return redirect(f'/dashboard/?month={current_month}&year={current_year}')

        students = Student.objects.filter(group_name__iexact=clean_group).order_by('full_name')
        if not students.exists():
            students = Student.objects.filter(group_name__icontains=clean_group).order_by('full_name')

        num_days = calendar.monthrange(current_year, current_month)[1]
        days_list = [{'num': d, 'is_sunday': date(current_year, current_month, d).weekday() == 6} for d in range(1, num_days + 1)]

        marks = {}
        all_atts = Attendance.objects.filter(date__month=current_month, date__year=current_year)
        for att in all_atts:
            day = att.date.day
            if att.status == 'invalid_absent': status_code = '4'
            elif att.status == 'valid_absent': status_code = '2'
            elif att.status == 'sick': status_code = '6'
            elif att.status == 'other': status_code = '8'
            elif att.status == 'present': status_code = '1'
            else: status_code = '0'
            marks[(att.student_id, day)] = {'status': status_code, 'color': att.color}

        group_table_data = []
        for student in students:
            is_this_headman = hasattr(student.user, 'headman_profile')
            row_cells = []
            total_missed = 0
            valid_missed = 0
            invalid_missed = 0
            
            for day in days_list:
                cell_data = marks.get((student.id, day['num']), {'status': '0', 'color': 'none'})
                row_cells.append({'day': day['num'], 'is_sunday': day['is_sunday'], 'status': cell_data['status'], 'color': cell_data['color']})
                
                color = cell_data['color']
                status = cell_data['status']
                if status in ['2', '4', '6', '8']:
                    hours = int(status)
                    if color == 'red':
                        invalid_missed += hours
                        total_missed += hours
                    elif color in ['blue', 'green', 'yellow']:
                        valid_missed += hours
                        total_missed += hours
                    
            group_table_data.append({
                'student': student, 'cells': row_cells, 'is_headman': is_this_headman,
                'total_missed': total_missed, 'valid_missed': valid_missed, 'invalid_missed': invalid_missed
            })

        return render(request, 'dashboard_curator.html', {
            'group_name': clean_group, 'role_title': role_title, 'students': group_table_data,
            'group_curator': Curator.objects.filter(group_name=clean_group).first(),
            'current_headman': Headman.objects.filter(group_name=clean_group).first(),
            'months': MONTHS, 'current_month': current_month, 'current_year': current_year,
            'years': range(2024, 2030), 'days': days_list
        })

    # --- 4. СТУДЕНТ ---
    elif is_student:
        student = request.user.student_profile
        student_name = student.full_name if student.full_name else request.user.username
        student_group = student.group_name if student.group_name else "Группа не указана"
        
        num_days = calendar.monthrange(current_year, current_month)[1]
        days_list = [{'num': d, 'is_sunday': date(current_year, current_month, d).weekday() == 6} for d in range(1, num_days + 1)]
        
        attendances = Attendance.objects.filter(student=student, date__month=current_month, date__year=current_year)
        
        marks = {}
        for att in attendances:
            day = att.date.day
            if att.status == 'invalid_absent': status_code = '4'
            elif att.status == 'valid_absent': status_code = '2'
            elif att.status == 'sick': status_code = '6'
            elif att.status == 'other': status_code = '8'
            elif att.status == 'present': status_code = '1'
            else: status_code = '0'
            marks[day] = {'status': status_code, 'color': att.color}

        row_cells = []
        total_missed, valid_missed, invalid_missed = 0, 0, 0
        
        for day in days_list:
            cell_data = marks.get(day['num'], {'status': '0', 'color': 'none'})
            row_cells.append({'day': day['num'], 'is_sunday': day['is_sunday'], 'status': cell_data['status'], 'color': cell_data['color']})
            
            color = cell_data['color']
            status = cell_data['status']
            if status in ['2', '4', '6', '8']:
                hours = int(status)
                if color == 'red':
                    invalid_missed += hours
                    total_missed += hours
                elif color in ['blue', 'green', 'yellow']:
                    valid_missed += hours
                    total_missed += hours

        return render(request, 'dashboard_student.html', {
            'student_name': student_name, 'student_group': student_group,
            'group_curator': Curator.objects.filter(group_name=student.group_name).first(),
            'days': days_list, 'cells': row_cells,
            'total_missed': total_missed, 'valid_missed': valid_missed, 'invalid_missed': invalid_missed,
            'months': MONTHS, 'current_month': current_month, 'current_year': current_year,
            'years': range(2024, 2030)
        })

    return redirect('login')


# --- СОХРАНЕНИЕ ---
def save_journal_marks(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            updates = data.get('updates', [])
            month = int(data.get('month'))
            year = int(data.get('year'))
            
            for item in updates:
                student_id = item['student_id']
                day = int(item['day'])
                status_code = item['status']
                color_code = item.get('color', 'none')
                
                date_str = f"{year}-{month:02d}-{day:02d}"
                
                if status_code == '0' and color_code == 'none':
                    Attendance.objects.filter(date=date_str, student_id=student_id).delete()
                else:
                    if status_code == '4': status_str = 'invalid_absent'
                    elif status_code == '2': status_str = 'valid_absent'
                    elif status_code == '6': status_str = 'sick'
                    elif status_code == '8': status_str = 'other'
                    elif status_code == '1': status_str = 'present'
                    elif status_code == '0': status_str = 'present'
                    else: continue 
                    
                    Attendance.objects.update_or_create(
                        date=date_str, 
                        student_id=student_id,
                        defaults={'status': status_str, 'color': color_code}
                    )
            
            return JsonResponse({'status': 'success'})
        except Exception as e:
            return JsonResponse({'status': 'error', 'message': str(e)})
    return JsonResponse({'status': 'error', 'message': 'Invalid request'})


# --- ЭКСПОРТ В EXCEL ---
def export_excel(request):
    current_month = int(request.GET.get('month', date.today().month))
    current_year = int(request.GET.get('year', date.today().year))
    group_name = request.GET.get('group', None)

    if group_name:
        students = Student.objects.filter(group_name=group_name).order_by('full_name')
    else:
        students = Student.objects.all().order_by('group_name', 'full_name')

    num_days = calendar.monthrange(current_year, current_month)[1]
    days_list = list(range(1, num_days + 1))

    marks = {}
    all_atts = Attendance.objects.filter(date__month=current_month, date__year=current_year)
    for att in all_atts:
        day = att.date.day
        if att.status == 'invalid_absent': status_code = '4'
        elif att.status == 'valid_absent': status_code = '2'
        elif att.status == 'sick': status_code = '6'
        elif att.status == 'other': status_code = '8'
        elif att.status == 'present': status_code = '1'
        else: status_code = '0'
        marks[(att.student_id, day)] = {'status': status_code, 'color': att.color}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Посещаемость"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(days_list) + 4)
    title_cell = ws.cell(row=1, column=1, value=f"Ведомость посещаемости за {MONTHS[current_month]} {current_year} г.")
    title_cell.font = Font(size=14, bold=True)
    title_cell.alignment = Alignment(horizontal='center')

    headers = ["№", "ФИО"] + [str(d) for d in days_list] + ["Пропущено", "Уваж.", "Неуваж."]
    ws.append(headers)

    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    row_num = 3
    for idx, student in enumerate(students, start=1):
        ws.cell(row=row_num, column=1, value=idx).alignment = Alignment(horizontal='center')
        ws.cell(row=row_num, column=2, value=student.full_name)

        total_missed, valid_missed, invalid_missed = 0, 0, 0

        for col_idx, day in enumerate(days_list, start=3):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

            cell_data = marks.get((student.id, day), {'status': '0', 'color': 'none'})
            status = cell_data['status']
            color = cell_data['color']

            if status in ['2', '4', '6', '8']:
                hours = int(status)
                if color == 'red':
                    cell.value = 'Н'
                    cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    cell.font = Font(color="9C0006")
                    invalid_missed += hours
                    total_missed += hours
                elif color == 'blue':
                    cell.value = 'У'
                    cell.fill = PatternFill(start_color="BDE0FE", end_color="BDE0FE", fill_type="solid")
                    cell.font = Font(color="000000")
                    valid_missed += hours
                    total_missed += hours
                elif color == 'green':
                    cell.value = 'Б'
                    cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                    cell.font = Font(color="006100")
                    valid_missed += hours
                    total_missed += hours
                elif color == 'yellow':
                    cell.value = 'Д'
                    cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                    cell.font = Font(color="9C6500")
                    valid_missed += hours
                    total_missed += hours
            elif status == '1':
                cell.value = '✔'
                cell.font = Font(color="006100")

        ws.cell(row=row_num, column=len(days_list) + 3, value=total_missed).border = thin_border
        ws.cell(row=row_num, column=len(days_list) + 4, value=valid_missed).border = thin_border
        ws.cell(row=row_num, column=len(days_list) + 5, value=invalid_missed).border = thin_border
        
        ws.cell(row=row_num, column=len(days_list) + 3).alignment = Alignment(horizontal='center')
        ws.cell(row=row_num, column=len(days_list) + 4).alignment = Alignment(horizontal='center')
        ws.cell(row=row_num, column=len(days_list) + 5).alignment = Alignment(horizontal='center')

        for col_idx in range(1, len(days_list) + 6):
            ws.cell(row=row_num, column=col_idx).border = thin_border

        row_num += 1

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 30
    for col_idx in range(3, len(days_list) + 3):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 5
    for col_idx in range(len(days_list) + 3, len(days_list) + 6):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 12

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="attendance_{group_name if group_name else "all"}_{current_month}_{current_year}.xlsx"'
    wb.save(response)
    return response