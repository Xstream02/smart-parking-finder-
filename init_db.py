from app import create_app, db
from app.models import User, ParkingSlot

app = create_app()

def initialize_database():
    with app.app_context():
        print("[DB Init] Creating database tables...")
        db.create_all()
        
        # Check if default admin exists, if not create
        admin_user = User.query.filter_by(username='admin').first()
        if not admin_user:
            print("[DB Init] Creating default admin user (admin / admin123)...")
            admin = User(username='admin', is_admin=True)
            admin.set_password('admin123')
            db.session.add(admin)
            
        # Check if default slots exist, if not create
        if ParkingSlot.query.count() == 0:
            print("[DB Init] Seeding default parking slots...")
            slots = [
                ParkingSlot(id=1, location='Block A - Spot 1 (Free)', status='Available', lat=18.5204, lng=73.8567, base_price=0.0, current_price=0.0),
                ParkingSlot(id=2, location='Block A - Spot 2 (Paid)', status='Available', lat=18.5210, lng=73.8575, base_price=100.0, current_price=100.0),
                ParkingSlot(id=3, location='Block B - Spot 1 (Free)', status='Available', lat=18.5215, lng=73.8580, base_price=0.0, current_price=0.0),
                ParkingSlot(id=4, location='Block B - Spot 2 (Paid)', status='Available', lat=18.5220, lng=73.8585, base_price=100.0, current_price=100.0),
                ParkingSlot(id=5, location='Library - Spot 1 (Paid)', status='Available', lat=18.5225, lng=73.8590, base_price=50.0, current_price=50.0),
                ParkingSlot(id=6, location='Library - Spot 2 (Free)', status='Available', lat=18.5230, lng=73.8595, base_price=0.0, current_price=0.0)
            ]
            for slot in slots:
                db.session.add(slot)
                
        db.session.commit()
        print("[DB Init] Database initialized and seeded successfully!")

if __name__ == '__main__':
    initialize_database()
