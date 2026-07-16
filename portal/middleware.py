from django.shortcuts import redirect
from django.contrib import messages
from django.urls import resolve

class GstRestrictedMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        if request.user.is_authenticated and request.user.username == 'gst':
            allowed_url_names = ['billing_records', 'logout', 'login', 'single_bill_pdf']
            try:
                resolver_match = resolve(request.path_info)
                url_name = resolver_match.url_name
            except Exception:
                url_name = None

            if url_name not in allowed_url_names:
                messages.error(request, "Access Denied: The GST user is restricted to Today's Billing Records only.")
                return redirect('billing_records')

        response = self.get_response(request)
        return response
