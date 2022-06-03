from django import forms

CHOICES = [(0, 'a'), (1, 'b'), (2, 'c')]

class UploadFileForm(forms.Form):
    file = forms.FileField(required=True)
    format = forms.ChoiceField(choices=CHOICES, widget=forms.RadioSelect, required=True)
