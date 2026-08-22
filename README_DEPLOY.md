# PathPilot - Final Render Deployment

## Render
1. Upload/push this entire folder to the GitHub repository used by Render.
2. Make sure `app.py`, `ml_engine.py`, `opencv_processor.py`, `templates/`, `static/`, and `requirements.txt` are in the repository root.
3. Render build command: `pip install -r requirements.txt`
4. Render start command: `gunicorn --bind 0.0.0.0:$PORT app:app`
5. Deploy.

The application initializes the SQLite schema when Gunicorn imports `app.py`, so the `users` table is created before registration/login requests.

Do not upload a local `data/learning.db` file. The app creates it automatically.
