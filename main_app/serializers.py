from rest_framework import serializers
from .models import TrainingRecord, TrainingAnswer

class TrainingAnswerSerializer(serializers.ModelSerializer):
    class Meta:
        model = TrainingAnswer
        fields = ['question_number', 'answer_value']

class TrainingRecordSerializer(serializers.ModelSerializer):
    answers = TrainingAnswerSerializer(many=True)

    class Meta:
        model = TrainingRecord
        fields = [
            'date', 'time', 'subject', 'student_name', "gender", "grade", "user_role",
            'school_operation_region', 'school_name', 
            'class_name', 'teacher_name', 'auto_correct_score_points', 'answers'
        ]

    def create(self, validated_data):
        answers_data = validated_data.pop('answers', [])
        training = TrainingRecord.objects.create(**validated_data)
        for ans in answers_data:
            TrainingAnswer.objects.create(training=training, **ans)
        return training
