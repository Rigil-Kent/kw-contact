from app import app, db
from app.models import Contact


@app.shell_context_processor
def get_shell_context():
    return {
        'db': db,
        'Contact': Contact
    }