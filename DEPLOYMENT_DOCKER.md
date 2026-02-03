# Docker Deployment Guide for Linux

This guide provides complete step-by-step instructions to deploy the S4Carlisle Hub Server application on a Linux server using Docker and Docker Compose.

## Table of Contents
1. [Prerequisites](#1-prerequisites)
2. [Initial Deployment](#2-initial-deployment)
3. [Database Migration (SQLite to PostgreSQL)](#3-database-migration-sqlite-to-postgresql)
4. [Admin User Setup](#4-admin-user-setup)
5. [SSL Certificate Setup](#5-ssl-certificate-setup)
6. [Service Management](#6-service-management)
7. [Maintenance](#7-maintenance)

---

## 1. Prerequisites

### Install Docker and Docker Compose

Ensure your Linux server has Docker and Docker Compose installed.

Follow the official guide for your distribution (e.g., Ubuntu):
https://docs.docker.com/engine/install/ubuntu/

**Quick install for Ubuntu:**
```bash
# Add Docker's official GPG key:
sudo apt-get update
sudo apt-get install ca-certificates curl gnupg
sudo install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
sudo chmod a+r /etc/apt/keyrings/docker.gpg

# Add the repository to Apt sources:
echo \
  "deb [arch=\"$(dpkg --print-architecture)\" signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu \
  $(. /etc/os-release && echo \"$VERSION_CODENAME\") stable" | \
  sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt-get update

# Install the Docker packages:
sudo apt-get install docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
```

**Verify Installation:**
```bash
sudo docker run hello-world
docker compose version
```

---

## 2. Initial Deployment

### 2.1 Clone Repository

Clone your repository to the server (e.g., `/var/www/hub-server`):

```bash
# Create directory and clone
sudo mkdir -p /var/www/hub-server
cd /var/www/hub-server/
sudo git clone https://github.com/Murali121290/my-project.git
cd my-project
```

Ensure the following files are present:
- `Dockerfile`
- `docker-compose.yml`
- `nginx/nginx.conf`
- `requirements.txt`

### 2.2 Set Permissions

Allow Docker to write to necessary directories:
```bash
sudo chmod -R 777 S4C-Processed-Documents reports logs
```

### 2.3 Start Services

Build and start all services:
```bash
sudo docker compose up -d --build
```

**Flags:**
- `--build`: Forces rebuild of images (required for first run)
- `-d`: Runs containers in background (detached mode)

### 2.4 Verify Deployment

Check that all containers are running:
```bash
sudo docker compose ps
```

You should see:
- `hub_server_web` - Flask application
- `hub_server_nginx` - Nginx reverse proxy
- `hub_postgres` - PostgreSQL database
- `hub_cms` - Directus CMS
- `certbot` - SSL certificate manager

View logs if needed:
```bash
sudo docker compose logs -f hub_web
```

---

## 3. Database Migration (SQLite to PostgreSQL)

If you have existing data in SQLite that needs to be migrated to PostgreSQL:

### 3.1 Export SQLite Data (on local machine)

```bash
# Run the export script
python export_db.py
```

This creates `postgres_dump.sql` with PostgreSQL-compatible INSERT statements.

### 3.2 Transfer to Server

Copy the dump file to your server:
```bash
scp postgres_dump.sql user@your-server:/var/www/hub-server/my-project/
```

### 3.3 Import to PostgreSQL

On the server, import the data:
```bash
cd /var/www/hub-server/my-project
sudo docker exec -i hub_postgres psql -U hub_user -d hub_db < postgres_dump.sql
```

**Note:** The dump file uses `ON CONFLICT DO NOTHING` to safely handle existing records.

---

## 4. Admin User Setup

### 4.1 Reset Admin Password

If you cannot log in with the default credentials (`admin` / `admin123`), reset the admin user:

```bash
cd /var/www/hub-server/my-project

# Copy the reset script into the container
sudo docker cp reset_admin.py hub_server_web:/app/reset_admin.py

# Run the reset script
sudo docker compose exec hub_web python reset_admin.py
```

You should see:
```
SUCCESS: Admin credentials updated.
Username: admin
Password: admin123
```

### 4.2 Login

Access your application at `http://your-server-ip/` and log in with:
- **Username:** `admin`
- **Password:** `admin123`

**Important:** Change the default password after first login!

---

## 5. SSL Certificate Setup

For production deployments, enable HTTPS with Let's Encrypt.

### 5.1 Prerequisites

- **Domain Name:** You must have a domain (e.g., `hub.example.com`)
- **DNS Configuration:** Domain's A record must point to your server's public IP
- **Ports:** 80 and 443 must be accessible from the internet

### 5.2 Update Configuration

**Edit `nginx/nginx.conf`** and replace all instances of `example.com` with your actual domain:

```bash
sudo nano nginx/nginx.conf
```

Find and replace (4 locations):
- `server_name example.com;` → `server_name hub.yourdomain.com;`
- `/etc/letsencrypt/live/example.com/` → `/etc/letsencrypt/live/hub.yourdomain.com/`

Save and commit changes:
```bash
git add nginx/nginx.conf
git commit -m "Update domain for SSL"
git push origin main
```

### 5.3 Run SSL Initialization

Make the script executable and run it:
```bash
chmod +x init-letsencrypt.sh
sudo ./init-letsencrypt.sh hub.yourdomain.com your-email@example.com
```

**Example:**
```bash
sudo ./init-letsencrypt.sh hub.s4carlisle.com admin@s4carlisle.com
```

The script will:
1. Create temporary dummy certificates
2. Start Nginx
3. Request real certificates from Let's Encrypt
4. Reload Nginx with new certificates

### 5.4 Verify HTTPS

Visit `https://hub.yourdomain.com` - you should see a valid SSL certificate!

**Certificate Auto-Renewal:** The `certbot` service automatically renews certificates every 12 hours.

---

## 6. Service Management

### Stop All Services
Stops and removes containers (data in volumes is preserved):
```bash
sudo docker compose down
```

### Start All Services
Starts all containers in background:
```bash
sudo docker compose up -d
```

Add `--build` if you've changed `Dockerfile` or `requirements.txt`:
```bash
sudo docker compose up -d --build
```

### Restart Specific Service
Restart only the web app (after code changes):
```bash
sudo docker compose restart hub_web
```

Restart Nginx (after config changes):
```bash
sudo docker compose restart nginx
```

### Check Status
View running containers:
```bash
sudo docker compose ps
```

### View Logs
View logs for a specific service:
```bash
sudo docker compose logs -f hub_web
sudo docker compose logs -f nginx
sudo docker compose logs -f postgres
```

---

## 7. Maintenance

### Updating Application Code

When you make changes to Python files or templates:

1. **Pull latest changes:**
   ```bash
   cd /var/www/hub-server/my-project
   git pull origin main
   ```

2. **Rebuild and restart:**
   ```bash
   sudo docker compose up -d --build hub_web
   ```

### Updating Nginx Configuration

1. **Edit configuration:**
   ```bash
   nano nginx/nginx.conf
   ```

2. **Test configuration:**
   ```bash
   sudo docker compose exec nginx nginx -t
   ```

3. **Reload Nginx:**
   ```bash
   sudo docker compose exec nginx nginx -s reload
   ```

### Database Backups

**Backup PostgreSQL:**
```bash
sudo docker exec hub_postgres pg_dump -U hub_user hub_db > backup_$(date +%Y%m%d).sql
```

**Restore from backup:**
```bash
sudo docker exec -i hub_postgres psql -U hub_user -d hub_db < backup_20260123.sql
```

### Accessing Directus CMS

Directus is available at `http://your-server-ip:8055/`

Default credentials (from `docker-compose.yml`):
- **Email:** `admin@example.com`
- **Password:** `admin`

---

## 8. Troubleshooting

### Application Not Accessible

1. **Check container status:**
   ```bash
   sudo docker compose ps
   ```

2. **Check logs:**
   ```bash
   sudo docker compose logs hub_web
   sudo docker compose logs nginx
   ```

3. **Verify ports are open:**
   ```bash
   sudo netstat -tulpn | grep -E '80|443'
   ```

### Database Connection Issues

1. **Check PostgreSQL is running:**
   ```bash
   sudo docker compose ps postgres
   ```

2. **Test database connection:**
   ```bash
   sudo docker exec -it hub_postgres psql -U hub_user -d hub_db
   ```

3. **Check environment variables:**
   ```bash
   sudo docker compose exec hub_web env | grep DB_
   ```

### SSL Certificate Issues

1. **Check Certbot logs:**
   ```bash
   sudo docker compose logs certbot
   ```

2. **Verify domain DNS:**
   ```bash
   nslookup hub.yourdomain.com
   ```

3. **Test certificate renewal:**
   ```bash
   sudo docker compose run --rm certbot renew --dry-run
   ```

### Permission Errors

If you see permission denied errors:
```bash
sudo chmod -R 777 S4C-Processed-Documents reports logs
sudo chown -R 1000:1000 certbot/
```

### Git SSL Verification Failed
If you see `server certificate verification failed` when pulling:
```bash
# Option 1: Update certificates (Recommended)
sudo apt-get update && sudo apt-get install --reinstall ca-certificates

# Option 2: Disable verification temporarily
git config http.sslVerify false
git pull origin main
git config http.sslVerify true
```

---

## 9. Important Notes

### Limitations on Linux
- **Word Automation:** Microsoft Word COM automation (`win32com`) does **not** work on Linux
- Features relying on macros will fail or return errors
- Python-native processing (regex, parsing) works fine

### Security Recommendations
1. Change default admin password immediately
2. Update Directus admin credentials
3. Use strong database passwords (update in `docker-compose.yml`)
4. Keep Docker and system packages updated
5. Configure firewall to only allow necessary ports (80, 443, 8055)

### Production Checklist
- [ ] Domain configured and DNS pointing to server
- [ ] SSL certificate installed and auto-renewal working
- [ ] Admin password changed from default
- [ ] Database backups configured
- [ ] Firewall rules configured
- [ ] Monitoring/logging set up
- [ ] Regular update schedule established