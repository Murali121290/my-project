# Linux Deployment Guide (Nginx + Gunicorn + SSL)

This guide details how to deploy the S4Carlisle Hub Server application on a Linux server (Ubuntu 20.04/22.04 recommended) using Nginx as a reverse proxy and Gunicorn as the application server, secured with a self-signed SSL certificate (OpenSSL).

## 1. Server Prerequisites

Update your system and install necessary packages:

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install python3-pip python3-venv nginx git openssl -y
```

## 2. Application Setup

1.  **Clone the Repository**
    We will deploy the application to `/var/www/hub-server`.

    ```bash
    sudo mkdir -p /var/www/hub-server
    sudo chown -R $USER:$USER /var/www/hub-server
    git clone https://github.com/Murali121290/my-project.git /var/www/hub-server
    cd /var/www/hub-server
    ```

2.  **Set up Python Environment**

    ```bash
    python3 -m venv venv
    source venv/bin/activate
    pip install --upgrade pip
    pip install -r requirements.txt
    pip install gunicorn
    ```

3.  **Initialize Database**
    Run the application once to create the database:
    ```bash
    export FLASK_APP=app_server.py
    flask run --port 5000
    # Press Ctrl+C to stop it after you see it start successfully
    ```

## 3. Configure Gunicorn Service

The repository includes a systemd service file (`hub_app.service`).

1.  **Edit the Service File**
    Open `hub_app.service` and ensure the `User`, `Group`, and paths match your server.
    
    ```bash
    nano hub_app.service
    ```
    *   Change `User=ubuntu` to your current linux username (check with `whoami`).
    *   Change `Group=ubuntu` to your user's group (usually same as username).
    *   Ensure `WorkingDirectory` is `/var/www/hub-server`.

2.  **Install and Start Service**

    ```bash
    sudo cp hub_app.service /etc/systemd/system/
    sudo systemctl daemon-reload
    sudo systemctl start hub_app
    sudo systemctl enable hub_app
    ```

3.  **Verify Status**
    ```bash
    sudo systemctl status hub_app
    ```
    (It should show "active (running)")

## 4. Generate SSL Certificate (OpenSSL)

For a private server or testing, generate a self-signed certificate.

1.  **Create Directory for Certificates**
    ```bash
    sudo mkdir -p /etc/nginx/ssl
    ```

2.  **Generate Certificate and Key**
    This command generates a certificate valid for 365 days.
    
    ```bash
    sudo openssl req -x509 -nodes -days 365 -newkey rsa:2048 \
      -keyout /etc/nginx/ssl/nginx-selfsigned.key \
      -out /etc/nginx/ssl/nginx-selfsigned.crt
    ```
    *   You will be asked for details (Country, State, Common Name, etc.). You can fill them in or press Enter to skip. For "Common Name", use your server's IP address or domain.

3.  **Generate Diffie-Hellman Group (Optional but recommended for security)**
    ```bash
    sudo openssl dhparam -out /etc/nginx/ssl/dhparam.pem 2048
    ```

## 5. Configure Nginx with SSL

1.  **Create Nginx Configuration**
    
    ```bash
    sudo nano /etc/nginx/sites-available/hub-app
    ```

2.  **Paste the Configuration**
    Replace `your_server_ip_or_domain` with your actual IP or domain name.

    ```nginx
    server {
        listen 80;
        server_name your_server_ip_or_domain;
        return 301 https://$host$request_uri;
    }

    server {
        listen 443 ssl;
        server_name your_server_ip_or_domain;

        ssl_certificate /etc/nginx/ssl/nginx-selfsigned.crt;
        ssl_certificate_key /etc/nginx/ssl/nginx-selfsigned.key;
        
        # SSL Settings
        ssl_protocols TLSv1.2 TLSv1.3;
        ssl_prefer_server_ciphers on;
        ssl_ciphers ECDHE-ECDSA-AES128-GCM-SHA256:ECDHE-RSA-AES128-GCM-SHA256:ECDHE-ECDSA-AES256-GCM-SHA384:ECDHE-RSA-AES256-GCM-SHA384:ECDHE-ECDSA-CHACHA20-POLY1305:ECDHE-RSA-CHACHA20-POLY1305:DHE-RSA-AES128-GCM-SHA256:DHE-RSA-AES256-GCM-SHA384;

        # App Proxy
        location / {
            include proxy_params;
            proxy_pass http://127.0.0.1:8000;
            proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
            proxy_set_header X-Forwarded-Proto $scheme;
            
            # Increase timeout for long processing tasks
            proxy_read_timeout 300;
            proxy_connect_timeout 300;
            proxy_send_timeout 300;
            
            # Increase upload size limit
            client_max_body_size 50M;
        }

        location /static {
            alias /var/www/hub-server/static;
        }
    }
    ```

3.  **Enable Configuration**
    ```bash
    sudo ln -s /etc/nginx/sites-available/hub-app /etc/nginx/sites-enabled/
    sudo rm /etc/nginx/sites-enabled/default  # Remove default welcome page
    ```

4.  **Test and Restart Nginx**
    ```bash
    sudo nginx -t
    sudo systemctl restart nginx
    ```

## 6. Permissions & Firewall

1.  **Fix Permissions**
    Ensure Nginx (running as `www-data` usually) and the App (running as your user) can access necessary files.
    
    ```bash
    # Allow uploads folder access
    chmod -R 775 /var/www/hub-server/S4C-Processed-Documents
    chmod -R 775 /var/www/hub-server/reports
    chmod -R 775 /var/www/hub-server/logs
    ```

2.  **Configure Firewall (UFW)**
    ```bash
    sudo ufw allow 'Nginx Full'
    sudo ufw enable
    ```

## 7. Accessing the Application

*   Open your browser and navigate to `https://your_server_ip`.
*   **Note**: Since you are using a self-signed certificate, your browser will show a security warning ("Your connection is not private"). Click **Advanced** -> **Proceed to... (unsafe)** to access the site.

---

### Production Note (Public Domain)
If you have a public domain name (e.g., `example.com`), it is recommended to use **Certbot** for a free, trusted SSL certificate instead of OpenSSL.

```bash
sudo apt install certbot python3-certbot-nginx
sudo certbot --nginx -d example.com
```
