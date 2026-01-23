#!/bin/bash

# Initialize Let's Encrypt SSL Certificate
# This script should be run ONCE on the server after deployment

if [ -z "$1" ]; then
    echo "Usage: ./init-letsencrypt.sh your-domain.com your-email@example.com"
    echo "Example: ./init-letsencrypt.sh hub.example.com admin@example.com"
    exit 1
fi

DOMAIN=$1
EMAIL=${2:-""}
STAGING=${3:-0} # Set to 1 for testing

echo "Initializing Let's Encrypt for domain: $DOMAIN"

# Create required directories
mkdir -p certbot/conf
mkdir -p certbot/www

# Download recommended TLS parameters
if [ ! -e "certbot/conf/options-ssl-nginx.conf" ] || [ ! -e "certbot/conf/ssl-dhparams.pem" ]; then
    echo "Downloading recommended TLS parameters..."
    curl -s https://raw.githubusercontent.com/certbot/certbot/master/certbot-nginx/certbot_nginx/_internal/tls_configs/options-ssl-nginx.conf > certbot/conf/options-ssl-nginx.conf
    curl -s https://raw.githubusercontent.com/certbot/certbot/master/certbot/certbot/ssl-dhparams.pem > certbot/conf/ssl-dhparams.pem
fi

echo "Creating dummy certificate for $DOMAIN..."
path="/etc/letsencrypt/live/$DOMAIN"
mkdir -p "certbot/conf/live/$DOMAIN"
docker compose run --rm --entrypoint "\
  openssl req -x509 -nodes -newkey rsa:4096 -days 1\
    -keyout '$path/privkey.pem' \
    -out '$path/fullchain.pem' \
    -subj '/CN=localhost'" certbot

echo "Starting nginx..."
docker compose up --force-recreate -d nginx

echo "Deleting dummy certificate for $DOMAIN..."
docker compose run --rm --entrypoint "\
  rm -Rf /etc/letsencrypt/live/$DOMAIN && \
  rm -Rf /etc/letsencrypt/archive/$DOMAIN && \
  rm -Rf /etc/letsencrypt/renewal/$DOMAIN.conf" certbot

echo "Requesting Let's Encrypt certificate for $DOMAIN..."

# Enable staging mode if needed
if [ $STAGING != "0" ]; then staging_arg="--staging"; fi

# Request certificate
if [ -z "$EMAIL" ]; then
    email_arg="--register-unsafely-without-email"
else
    email_arg="--email $EMAIL"
fi

docker compose run --rm --entrypoint "\
  certbot certonly --webroot -w /var/www/certbot \
    $staging_arg \
    $email_arg \
    -d $DOMAIN \
    --rsa-key-size 4096 \
    --agree-tos \
    --force-renewal" certbot

echo "Reloading nginx..."
docker compose exec nginx nginx -s reload

echo "Done! Your site should now be accessible via HTTPS at https://$DOMAIN"
