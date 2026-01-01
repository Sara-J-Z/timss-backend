from django.http import HttpResponse

from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status

from .serializers import TrainingRecordSerializer
from .excel_utils import save_to_excel


class SubmitTrainingAPIView(APIView):
    """
    Resilient endpoint:
    - Try to save in DB (if DB is available)
    - Always try to update/upload Excel to OneDrive
    - If DB is down, system still works using OneDrive as the source of truth
    """

    def post(self, request, *args, **kwargs):
        training_id = None
        db_saved = False
        db_error = None

        # 1) Validate payload structure (serializer validation)
        serializer = TrainingRecordSerializer(data=request.data)
        if serializer.is_valid():
            # 2) Try saving to DB (may fail if DB is down)
            try:
                training = serializer.save()
                training_id = training.id
                db_saved = True
            except Exception as e:
                db_saved = False
                db_error = str(e)
        else:
            # If invalid, we still attempt Excel save (optional but useful)
            db_error = serializer.errors

        # 3) Always attempt Excel + OneDrive update (source of truth)
        excel_saved = False
        excel_error = None
        try:
            save_to_excel(request.data)  # use raw payload, no DB dependency
            excel_saved = True
        except Exception as e:
            excel_saved = False
            excel_error = str(e)

        # 4) Decide response status
        # If Excel failed, that's the critical failure for your workflow
        if not excel_saved:
            return Response({
                "message": "Processed, but Excel/OneDrive failed",
                "db_saved": db_saved,
                "training_id": training_id,
                "db_error": db_error,
                "excel_saved": excel_saved,
                "excel_error": excel_error,
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

        # If Excel succeeded, return success even if DB failed
        return Response({
            "message": "Processed successfully",
            "db_saved": db_saved,
            "training_id": training_id,
            "db_error": db_error,
            "excel_saved": excel_saved,
            "excel_error": excel_error,
        }, status=status.HTTP_201_CREATED)


def azure_callback(request):
    """
    Dummy endpoint for Azure App Registration.
    This endpoint is NOT used in authentication flow.
    """
    return HttpResponse("Azure callback endpoint is configured.")
