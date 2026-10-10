import re

from django.db import transaction
from rest_framework import serializers

from .models import TrainingRecord, TrainingAnswer


_INVALID_PATH_CHARS = re.compile(r'[\\/:*?"<>|]')


class TrainingAnswerSerializer(serializers.ModelSerializer):
    class Meta:
        model = TrainingAnswer
        fields = ['question_number', 'answer_value']


class TrainingRecordSerializer(serializers.ModelSerializer):
    answers = TrainingAnswerSerializer(many=True)

    class Meta:
        model = TrainingRecord
        fields = [
            'date', 'time', 'subject', 'student_name', 'gender', 'grade', 'user_role',
            'school_operation_region', 'school_name',
            'class_name', 'teacher_name', 'auto_correct_score_points', 'answers'
        ]

    def validate_school_name(self, value):
        """Reject values that are unsafe to use as a OneDrive folder/file name."""
        name = (value or '').strip()

        if not name:
            raise serializers.ValidationError('school_name is required.')

        # Prevent path traversal / special path segments.
        if name in {'.', '..'} or '..' in name:
            raise serializers.ValidationError('Invalid school_name.')

        if _INVALID_PATH_CHARS.search(name):
            raise serializers.ValidationError('school_name contains invalid characters.')

        # OneDrive/SharePoint file and folder names should not end with a dot/space.
        if name.endswith('.') or name.endswith(' '):
            raise serializers.ValidationError('Invalid school_name ending.')

        return name

    @transaction.atomic
    def create(self, validated_data):
        answers_data = validated_data.pop('answers', [])
        training = TrainingRecord.objects.create(**validated_data)
        for ans in answers_data:
            TrainingAnswer.objects.create(training=training, **ans)
        return training
