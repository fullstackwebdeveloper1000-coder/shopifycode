import os, sys, http.server, socketserver
os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))
PORT = 4322
socketserver.TCPServer.allow_reuse_address = True
with socketserver.TCPServer(("127.0.0.1", PORT), http.server.SimpleHTTPRequestHandler) as httpd:
    print(f"serving site on http://127.0.0.1:{PORT}"); httpd.serve_forever()
