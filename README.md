# Salary App - Payslip Generation

A Flask-based payslip generation application with authentication and Excel data integration.

## Features
- User authentication with bcrypt
- Employee data loading from Excel
- Payslip generation (DOCX format)
- Responsive web interface

## Local Development

### Prerequisites
- Python 3.8+
- pip

### Setup
1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Create a user account:
```bash
python create_user.py
```

3. Run the development server:
```bash
python app.py
```

The app will be available at `http://localhost:5000`

## Deployment on Render

### Prerequisites
- A GitHub repository with this code
- A Render account (free tier available)

### Steps
1. Push your code to GitHub
2. Go to [render.com](https://render.com)
3. Create a new **Web Service**
4. Connect your GitHub repository
5. Render will automatically detect the `render.yaml` configuration
6. Set the following environment variables in Render dashboard:
   - `SECRET_KEY` - A secure random string (auto-generated if using render.yaml)
   - `FLASK_ENV` - Set to `production`
   - `EXCEL_PATH` - (optional) Path to Excel file if needed

7. Deploy!

### Post-Deployment
After deployment, you'll need to:
1. Add a user via SSH or direct database access (since create_user.py creates users.json locally)
2. Upload your Excel file with employee data if not already in the repo

## Files
- `app.py` - Main Flask application
- `requirements.txt` - Python dependencies
- `Procfile` - Render/Heroku process configuration
- `render.yaml` - Render-specific deployment configuration
- `create_user.py` - User account creation utility
- `generate_payslip.py` - Payslip DOCX generation
- `templates/` - HTML templates

## Environment Variables
- `PORT` - Server port (default: 5000)
- `SECRET_KEY` - Flask session secret key
- `EXCEL_PATH` - Path to employee Excel file
- `FLASK_ENV` - `development` or `production`
