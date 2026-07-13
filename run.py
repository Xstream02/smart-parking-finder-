import os
from app import create_app, socketio
from init_db import initialize_database

app = create_app()

if __name__ == '__main__':
    # Auto-initialize database if not present
    if not os.path.exists('database.db'):
        initialize_database()
        
    print("--------------------------------------------------")
    # Using eventlet or standard gevent or development threading
    print("   Starting Production-Ready Smart Parking Finder  ")
    print("   Port: 5555 | Socket.IO Active                   ")
    print("--------------------------------------------------")
    socketio.run(app, debug=True, use_reloader=False, port=5555)
