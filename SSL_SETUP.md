# SSL Certificate Setup Guide

## Prerequisites
1. **Domain Name**: You must have a domain (e.g., `hub.example.com`) pointed to your server's IP address
2. **DNS Configuration**: Ensure your domain's A record points to your server's public IP
3. **Ports Open**: Ports 80 and 443 must be accessible from the internet

## Setup Steps

### 1. Update Configuration Files

**Edit `nginx/nginx.conf`** and replace all instances of `example.com` with your actual domain:
```bash
# Replace example.com with your domain in these lines:
server_name example.com;  # Change to: server_name hub.yourdomain.com;
ssl_certificate /etc/letsencrypt/live/example.com/fullchain.pem;
ssl_certificate_key /etc/letsencrypt/live/example.com/privkey.pem;
```

### 2. Push Changes to Repository
```bash
git add docker-compose.yml nginx/nginx.conf init-letsencrypt.sh
git commit -m "Configure Let's Encrypt SSL"
git push origin main
```

### 3. On Your Server

Pull the latest changes:
```bash
cd /var/www/hub-server/my-project
git pull origin main
```

Make the initialization script executable:
```bash
chmod +x init-letsencrypt.sh
```

Run the Let's Encrypt initialization script:
```bash
sudo ./init-letsencrypt.sh your-domain.com your-email@example.com
```

**Example:**
```bash
sudo ./init-letsencrypt.sh hub.example.com admin@example.com
```

The script will:
- Create a temporary dummy certificate
- Start Nginx
- Request a real certificate from Let's Encrypt
- Reload Nginx with the new certificate

### 4. Verify

Visit your site at `https://your-domain.com` - you should see a valid SSL certificate!

## Certificate Renewal

The `certbot` service in `docker-compose.yml` automatically renews certificates every 12 hours. No manual intervention needed.

## Troubleshooting

### "Connection Refused" or "Cannot Connect"
- Verify your domain's DNS A record points to your server IP
- Check that ports 80 and 443 are open in your firewall
- Verify Docker containers are running: `sudo docker compose ps`

### "Failed Authorization Procedure"
- Ensure your domain is correctly configured in `nginx/nginx.conf`
- Check Nginx logs: `sudo docker compose logs nginx`
- Verify the `.well-known/acme-challenge/` location is accessible

### Testing Mode
To test without hitting Let's Encrypt rate limits, use staging mode:
```bash
sudo ./init-letsencrypt.sh your-domain.com your-email@example.com 1
```

## Using DNS API Key (Advanced)

If you have a DNS provider API key (e.g., Cloudflare, Route53), you can use DNS-01 challenge instead of HTTP-01. This allows:
- Wildcard certificates (*.example.com)
- Certificate generation without opening port 80

This requires additional Certbot plugins. Let me know if you need this setup.
