from django import forms
from django.core.exceptions import ValidationError

class PDFUploadForm(forms.Form):
    pdf_file = forms.FileField(
        label="Upload PDF",
        required=True,
        widget=forms.ClearableFileInput(attrs={"accept": ".pdf"})
    )

    def clean_pdf_file(self):
        pdf = self.cleaned_data.get("pdf_file")
        if pdf.size > 10 * 1024 * 1024:
            raise ValidationError("Max file size is 10MB")
        return pdf
