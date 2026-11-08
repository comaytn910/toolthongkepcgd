import socket
import time

s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.bind(('127.0.0.1', 8501))
s.listen(1)
print("FAKE_SERVER_LISTENING_8501", flush=True)
time.sleep(25)
s.close()
