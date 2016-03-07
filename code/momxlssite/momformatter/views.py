from django.shortcuts import render
from django.http import HttpResponseRedirect, HttpResponse

from momformatter.forms import UploadFileForm
from momformatter.formatter import formatter


def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        print("Get the form!")
        if form.is_valid():
            _handle_uploaded_file(request.FILES['file'])

            #return HttpResponseRedirect('/successful/url/')
            return HttpResponse("Hello World")
    else:
        form = UploadFileForm()

    return render(request, 'upload.html', {'form': form})


def _handle_uploaded_file(f):
    # TODO: Create a temp folder to keep these temp files.
    temp_input_file = 'temp_input.xlsx'
    temp_output_file = 'temp_output.csv'
    with open('temp_input.xlsx', 'wb+') as destination:
        for chunk in f.chunks():
            destination.write(chunk)

    formatter.format_xls(temp_input_file, temp_output_file, sheet_index=1, name_col=0, address_col=2)
