from django.shortcuts import render
from django.http import HttpResponse
import csv

from momformatter.forms import UploadFileForm
from momformatter.formatter import formatter


def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            lines = _format_uploaded_file_into_writable_list(request.FILES['file'])

            response = _create_csv_response(lines)
            return response
    else:
        form = UploadFileForm()

    return render(request, 'upload.html', {'form': form})


def _create_csv_response(lines):
    # Create the HttpResponse object with the appropriate CSV header.
    response = HttpResponse(content_type='text/csv; charset=utf-16')
    response['Content-Disposition'] = 'attachment; filename="result.csv"'

    writer = csv.writer(response, delimiter='\t')

    for line in lines:
        writer.writerow(line)

    return response


def _format_uploaded_file_into_writable_list(f):
    # TODO: Create a temp folder to keep these temp files.
    temp_input_file = 'temp_input.xlsx'
    with open('temp_input.xlsx', 'wb+') as destination:
        for chunk in f.chunks():
            destination.write(chunk)

    lines = formatter.format_xls_to_csv_list(temp_input_file)

    return lines
