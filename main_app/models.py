import datetime
from django.db import models

class TrainingRecord(models.Model):
    # معلومات التدريب
    date = models.DateField(default=datetime.date.today)  # default لتجنب مشاكل الصفوف القديمة
    time = models.TimeField(default=datetime.datetime.now)
    subject = models.CharField(max_length=100, default='Unknown')
    
    # بيانات الطالب والمدرسة
    student_name = models.CharField(max_length=200, default='Unknown')
    school_operation_region = models.CharField(max_length=200, default='Unknown')
    school_name = models.CharField(max_length=200, default='Unknown')
    class_name = models.CharField(max_length=100, default='Unknown')
    teacher_name = models.CharField(max_length=200, default='Unknown')
    
    # مجموع النقاط (من Storyline)
    auto_correct_score_points = models.IntegerField(null=True, blank=True)
    
    # تاريخ ووقت إنشاء السجل
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.student_name} - {self.subject}"


class TrainingAnswer(models.Model):
    training = models.ForeignKey(
        TrainingRecord,
        on_delete=models.CASCADE,
        related_name="answers"
    )
    question_number = models.CharField(max_length=10)   # مثال: "Q1", "Q1a"
    answer_value = models.CharField(max_length=500, blank=True, null=True)

    def __str__(self):
        return f"{self.training.student_name} - {self.question_number}"
