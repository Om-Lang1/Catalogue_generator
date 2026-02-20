import pypyodbc as odbc
from flask import Flask, render_template_string

# Database connection configuration
DRIVER_NAME = 'SQL SERVER'
SERVER_NAME = 'SJSERVER'
DATABASE_NAME = 'TSHISJEP'

def test_db_connection():
    """Test database connection and return status"""
    try:
        connection_string = f"""
            DRIVER={{{DRIVER_NAME}}};
            SERVER={SERVER_NAME};
            DATABASE={DATABASE_NAME};
            Trust_Connection=yes;
        """
        conn = odbc.connect(connection_string)
        conn.close()
        return True, "Database connected successfully!"
    except Exception as e:
        return False, f"Database connection failed: {str(e)}"

app = Flask(__name__)

@app.route('/')
def index():
    """Database connection status page"""
    is_connected, message = test_db_connection()
    
    # Simple HTML template
    html_template = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Database Connection Status</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                max-width: 600px;
                margin: 50px auto;
                padding: 20px;
                background-color: #f5f5f5;
            }
            .container {
                background-color: white;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                text-align: center;
            }
            .status {
                font-size: 24px;
                font-weight: bold;
                margin: 20px 0;
                padding: 15px;
                border-radius: 5px;
            }
            .connected {
                background-color: #d4edda;
                color: #155724;
                border: 1px solid #c3e6cb;
            }
            .disconnected {
                background-color: #f8d7da;
                color: #721c24;
                border: 1px solid #f5c6cb;
            }
            .details {
                margin-top: 20px;
                padding: 15px;
                background-color: #f8f9fa;
                border-radius: 5px;
                text-align: left;
            }
            .refresh-btn {
                background-color: #007bff;
                color: white;
                padding: 10px 20px;
                border: none;
                border-radius: 5px;
                cursor: pointer;
                font-size: 16px;
                margin-top: 20px;
            }
            .refresh-btn:hover {
                background-color: #0056b3;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Database Connection Status</h1>
            
            <div class="status {{ status_class }}">
                {{ status_icon }} {{ status_text }}
            </div>
            
            <div class="details">
                <h3>Connection Details:</h3>
                <p><strong>Server:</strong> {{ server_name }}</p>
                <p><strong>Database:</strong> {{ database_name }}</p>
                <p><strong>Driver:</strong> {{ driver_name }}</p>
                <p><strong>Status Message:</strong> {{ message }}</p>
            </div>
            
            <button class="refresh-btn" onclick="location.reload()">Refresh Status</button>
        </div>
        
        <script>
            // Auto-refresh every 30 seconds
            setTimeout(function(){
                location.reload();
            }, 30000);
        </script>
    </body>
    </html>
    """
    
    return render_template_string(
        html_template,
        status_class='connected' if is_connected else 'disconnected',
        status_icon='✅' if is_connected else '❌',
        status_text='Connected' if is_connected else 'Disconnected',
        server_name=SERVER_NAME,
        database_name=DATABASE_NAME,
        driver_name=DRIVER_NAME,
        message=message
    )

if __name__ == '__main__':
    print("Starting Database Connection Checker...")
    print(f"Testing connection to {SERVER_NAME}\\{DATABASE_NAME}")
    
    # Test connection on startup
    is_connected, message = test_db_connection()
    print(f"Initial test: {message}")
    
    # Start Flask app
    app.run(debug=True, host='0.0.0.0', port=5000)
