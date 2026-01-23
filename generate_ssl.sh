#!/bin/bash

# Create ssl directory if it doesn't exist
mkdir -p nginx/ssl

# Generate Self-Signed Certificate
# -nodes: No password
# -days 365: Valid for a year
# -newkey rsa:2048: 2048 bit RSA key
# -keyout: Output key file
# -out: Output certificate file
# -subj: Subject (avoid interactive prompt)
echo "Generating Self-Signed SSL Certificate..."

openssl req -x509 -nodes -days 365 -newkey rsa:2048 \
    -keyout nginx/ssl/nginx.key \
    -out nginx/ssl/nginx.crt \
    -subj "/C=US/ST=State/L=City/O=Organization/CN=localhost"

echo "Certificate created in nginx/ssl/"
chmod 644 nginx/ssl/nginx.crt
chmod 600 nginx/ssl/nginx.key
