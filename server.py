import http.server, os

os.chdir("/Users/mariagrilo/Documents/AlcoveIQ")

class Handler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # suppress access logs

httpd = http.server.HTTPServer(("", 3000), Handler)
httpd.serve_forever()
