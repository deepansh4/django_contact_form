from rest_framework import serializers
from .models import ContactSubmission

class ContactSubmissionSerializer(serializers.ModelSerializer):
    class Meta:
        model = ContactSubmission
        fields = [
            'id',
            'full_name',
            'email',
            'mobile_number',
            'category',
            'sub_category',
            'agreed_to_terms',
            'created_at'
        ]
        read_only_fields = ['id', 'created_at']

    def validate_agreed_to_terms(self, value):
        if not value:
            raise serializers.ValidationError("You must agree to the terms and conditions.")
        return value

    def validate_email(self, value):
        if ContactSubmission.objects.filter(email=value).exists():
            raise serializers.ValidationError("A submission with this email already exists.")
        return value
