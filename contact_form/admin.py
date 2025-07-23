from django.contrib import admin
from django.http import HttpResponse
from django.urls import path, reverse
from django.shortcuts import redirect
from django.contrib import messages
from django.utils.html import format_html
import pandas as pd
from io import BytesIO
from datetime import datetime
from .models import ContactSubmission


def export_selected_to_excel(modeladmin, request, queryset):
    """Export selected contact submissions to Excel"""
    try:
        # Convert queryset to list of dictionaries
        data = []
        for submission in queryset:
            data.append({
                'ID': submission.id,
                'Full Name': submission.full_name,
                'Email': submission.email,
                'Mobile Number': submission.mobile_number,
                'Category': submission.category,
                'Sub Category': submission.sub_category,
                'Agreed to Terms': 'Yes' if submission.agreed_to_terms else 'No',
                'Created At': submission.created_at.replace(
                    tzinfo=None) if submission.created_at.tzinfo else submission.created_at
            })

        # Create DataFrame
        df = pd.DataFrame(data)

        if df.empty:
            modeladmin.message_user(request, 'No data selected to export', level='WARNING')
            return None

        # Create Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Contact Submissions', index=False)

            # Auto-adjust column widths
            workbook = writer.book
            worksheet = writer.sheets['Contact Submissions']

            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        output.seek(0)

        # Create response
        response = HttpResponse(
            output.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

        filename = f'selected_contact_submissions_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        response['Content-Disposition'] = f'attachment; filename="{filename}"'

        # Success message
        modeladmin.message_user(request, f'Successfully exported {len(data)} submissions to Excel', level='SUCCESS')

        return response

    except Exception as e:
        modeladmin.message_user(request, f'Export failed: {str(e)}', level='ERROR')
        return None


export_selected_to_excel.short_description = "📊 Export selected submissions to Excel"


def export_all_to_excel(modeladmin, request, queryset):
    """Export ALL contact submissions to Excel (ignores selection)"""
    try:
        # Get all submissions
        all_submissions = ContactSubmission.objects.all()

        # Convert to list of dictionaries
        data = []
        for submission in all_submissions:
            data.append({
                'ID': submission.id,
                'Full Name': submission.full_name,
                'Email': submission.email,
                'Mobile Number': submission.mobile_number,
                'Category': submission.category,
                'Sub Category': submission.sub_category,
                'Agreed to Terms': 'Yes' if submission.agreed_to_terms else 'No',
                'Created At': submission.created_at.replace(
                    tzinfo=None) if submission.created_at.tzinfo else submission.created_at
            })

        # Create DataFrame
        df = pd.DataFrame(data)

        if df.empty:
            modeladmin.message_user(request, 'No data available to export', level='WARNING')
            return None

        # Create Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='All Contact Submissions', index=False)

            # Auto-adjust column widths
            workbook = writer.book
            worksheet = writer.sheets['All Contact Submissions']

            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        output.seek(0)

        # Create response
        response = HttpResponse(
            output.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

        filename = f'all_contact_submissions_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        response['Content-Disposition'] = f'attachment; filename="{filename}"'

        # Success message
        modeladmin.message_user(request, f'Successfully exported all {len(data)} submissions to Excel', level='SUCCESS')

        return response

    except Exception as e:
        modeladmin.message_user(request, f'Export failed: {str(e)}', level='ERROR')
        return None


export_all_to_excel.short_description = "📊 Export ALL submissions to Excel"


@admin.register(ContactSubmission)
class ContactSubmissionAdmin(admin.ModelAdmin):
    list_display = [
        'full_name',
        'email',
        'mobile_number',
        'category',
        'sub_category',
        'agreed_to_terms',
        'created_at'
    ]
    list_filter = ['category', 'sub_category', 'agreed_to_terms', 'created_at']
    search_fields = ['full_name', 'email', 'mobile_number']
    readonly_fields = ['created_at', 'updated_at']
    ordering = ['-created_at']

    # Add both export actions
    actions = [export_selected_to_excel, export_all_to_excel]

    def get_urls(self):
        """Add custom URLs for additional export functionality"""
        urls = super().get_urls()
        custom_urls = [
            path('export-all-excel/', self.export_all_excel_view, name='contact_export_all_excel'),
        ]
        return custom_urls + urls

    def export_all_excel_view(self, request):
        """Custom view for exporting all submissions with a direct URL"""
        try:
            all_submissions = ContactSubmission.objects.all()

            # Convert to list of dictionaries
            data = []
            for submission in all_submissions:
                data.append({
                    'ID': submission.id,
                    'Full Name': submission.full_name,
                    'Email': submission.email,
                    'Mobile Number': submission.mobile_number,
                    'Category': submission.category,
                    'Sub Category': submission.sub_category,
                    'Agreed to Terms': 'Yes' if submission.agreed_to_terms else 'No',
                    'Created At': submission.created_at.replace(
                        tzinfo=None) if submission.created_at.tzinfo else submission.created_at
                })

            # Create DataFrame
            df = pd.DataFrame(data)

            if df.empty:
                messages.warning(request, 'No data available to export')
                return redirect('admin:contact_form_contactsubmission_changelist')

            # Create Excel file
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='All Contact Submissions', index=False)

                # Auto-adjust column widths
                workbook = writer.book
                worksheet = writer.sheets['All Contact Submissions']

                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            output.seek(0)

            # Create response
            response = HttpResponse(
                output.getvalue(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

            filename = f'all_contact_submissions_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
            response['Content-Disposition'] = f'attachment; filename="{filename}"'

            return response

        except Exception as e:
            messages.error(request, f'Export failed: {str(e)}')
            return redirect('admin:contact_form_contactsubmission_changelist')

    def changelist_view(self, request, extra_context=None):
        """Add extra context for custom export button"""
        extra_context = extra_context or {}
        extra_context['custom_export_url'] = reverse('admin:contact_export_all_excel')
        return super().changelist_view(request, extra_context)
