def hostname():
    import socket
    hostname = socket.gethostname()
    return hostname