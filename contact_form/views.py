from rest_framework import status

from rest_framework.response import Response
from django.views.decorators.csrf import csrf_exempt
from .serializers import ContactSubmissionSerializer
import pandas as pd
from django.http import HttpResponse
from io import BytesIO
from rest_framework.decorators import api_view
from .models import ContactSubmission
import pandas as pd
from django.http import HttpResponse
from django.utils import timezone
from io import BytesIO
from datetime import datetime
from rest_framework.decorators import api_view
from rest_framework.response import Response
from .models import ContactSubmission


@api_view(['POST'])
@csrf_exempt
def submit_contact_form(request):
    """
    Handle contact form submission
    """
    try:
        serializer = ContactSubmissionSerializer(data=request.data)

        if serializer.is_valid():
            contact_submission = serializer.save()

            return Response({
                'success': True,
                'message': 'Contact form submitted successfully!',
                'data': {
                    'id': contact_submission.id,
                    'full_name': contact_submission.full_name,
                    'email': contact_submission.email,
                    'created_at': contact_submission.created_at
                }
            }, status=status.HTTP_201_CREATED)

        return Response({
            'success': False,
            'message': 'Form validation failed',
            'errors': serializer.errors
        }, status=status.HTTP_400_BAD_REQUEST)

    except Exception as e:
        return Response({
            'success': False,
            'message': 'An error occurred while processing your request',
            'error': str(e)
        }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


@api_view(['GET'])
def get_contact_submissions(request):
    """
    Get all contact submissions (for admin purposes)
    """
    submissions = ContactSubmission.objects.all()
    serializer = ContactSubmissionSerializer(submissions, many=True)

    return Response({
        'success': True,
        'data': serializer.data,
        'count': submissions.count()
    })


# @api_view(['GET'])
# def export_contact_submissions_to_excel(request):
#     """
#     Export all contact submissions to Excel file
#     """
#     try:
#         # Query all contact submissions
#         queryset = ContactSubmission.objects.all()
#
#         # Convert to list of dictionaries with timezone-unaware datetimes
#         data = []
#         for submission in queryset:
#             data.append({
#                 'ID': submission.id,
#                 'Full Name': submission.full_name,
#                 'Email': submission.email,
#                 'Mobile Number': submission.mobile_number,
#                 'Category': submission.category,
#                 'Sub Category': submission.sub_category,
#                 'Agreed to Terms': 'Yes' if submission.agreed_to_terms else 'No',
#                 'Created At': submission.created_at.replace(
#                     tzinfo=None) if submission.created_at.tzinfo else submission.created_at
#             })
#
#         # Convert to DataFrame
#         df = pd.DataFrame(data)
#
#         if df.empty:
#             return Response({
#                 'success': False,
#                 'message': 'No data available to export'
#             }, status=404)
#
#         # Create a BytesIO buffer to hold the Excel data
#         output = BytesIO()
#
#         # Write DataFrame to Excel with formatting
#         with pd.ExcelWriter(output, engine='openpyxl') as writer:
#             df.to_excel(writer, sheet_name='Contact Submissions', index=False)
#
#             # Get the workbook and worksheet for formatting
#             workbook = writer.book
#             worksheet = writer.sheets['Contact Submissions']
#
#             # Auto-adjust column widths
#             for column in worksheet.columns:
#                 max_length = 0
#                 column_letter = column[0].column_letter
#                 for cell in column:
#                     try:
#                         if len(str(cell.value)) > max_length:
#                             max_length = len(str(cell.value))
#                     except:
#                         pass
#                 adjusted_width = min(max_length + 2, 50)
#                 worksheet.column_dimensions[column_letter].width = adjusted_width
#
#         # Rewind the buffer
#         output.seek(0)
#
#         # Create HTTP response
#         response = HttpResponse(
#             output.getvalue(),
#             content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
#         )
#
#         # Set filename with current timestamp
#         filename = f'contact_submissions_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
#         response['Content-Disposition'] = f'attachment; filename={filename}'
#
#         return response
#
#     except Exception as e:
#         return Response({
#             'success': False,
#             'message': 'Error exporting data to Excel',
#             'error': str(e)
#         }, status=500)
#
