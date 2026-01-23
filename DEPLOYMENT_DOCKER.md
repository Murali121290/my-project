# Docker Deployment Guide for Linux

This guide details how to deploy the S4Carlisle Hub Server application on a Linux server using Docker and Docker Compose. This simplifies the setup by containerizing the application, database (Postgres for Directus), and web server (Nginx).

## 1. Prerequisites

Ensure your Linux server has Docker and Docker Compose installed.

1.  **Install Docker Engine & Docker Compose**:
    Follow the official guide for your distribution (e.g., Ubuntu):
    https://docs.docker.com/engine/install/ubuntu/

    Quick install for Ubuntu:
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

2.  **Verify Installation**:
    ```bash
    sudo docker run hello-world
    docker compose version
    ```

## 2. Deploying the Application

1.  **Transfer Files to Server**:
    Clone your repository or copy your project files to the server (e.g., `/var/www/hub-server`).

    ```bash
    # Example using git
    git clone https://github.com/Murali121290/my-project.git /var/www/hub-server
    cd /var/www/hub-server
    ```

    Ensure the following new files are present:
    -   `Dockerfile`
    -   `docker-compose.yml`
    -   `nginx/nginx.conf`
    -   `requirements.txt`

2.  **Prepare SQLite Database File**:
    The application uses `reference_validator.db`. If it doesn't exist yet, create an empty file so Docker can mount it correctly:
    ```bash
    touch reference_validator.db
    ```
    *Note: If you already have data, copy your local `reference_validator.db` to the server.*

3.  **Start the Services**:
    Run the following command in the project directory:
    ```bash
    sudo docker compose up -d --build
    ```
    -   `--build`: Forces a rebuild of the Python image (useful for first run or updates).
    -   `-d`: Runs containers in detached mode (background).

4.  **Check Status**:
    ```bash
    sudo docker compose ps
    ```
    You should see `hub_server_web`, `hub_server_nginx`, `hub_postgres`, and `hub_cms` running.

5.  **View Logs** (if needed):
    ```bash
    sudo docker compose logs -f hub_web
    ```

## 3. Configuration & Maintenance

### Permissions
Since Docker containers run as root (or a specific user), ensure the mounted volumes on the host are accessible.
```bash
# Allow the container to write to these directories
sudo chmod -R 777 S4C-Processed-Documents reports logs
sudo chmod 666 reference_validator.db
```

### Updating the Application
When you change code (Python files, templates):
1.  Pull new changes: `git pull`
2.  Rebuild and restart:
    ```bash
    sudo docker compose up -d --build hub_web
    ```

### Accessing the App
-   **Web App**: `http://<your-server-ip>/`
-   **Directus CMS**: `http://<your-server-ip>:8055/`

### Directus (CMS) Setup
The first time `directus` runs, it will initialize the Postgres database.
-   Login to Directus at port 8055 using the credentials in `docker-compose.yml`:
    -   Email: `admin@example.com`
    -   Password: `admin`

## 4. Important Limitations on Linux
-   **Macros/Word Automation**: As noted in the app config, Microsoft Word COM automation does **not** work on Linux. Features relying on `win32com` (like macro processing) will fail or return errors.
-   **File Processing**: Python-native processing (regex, parsing) will work fine.

## 5. SSL (HTTPS)
To use HTTPS (installing a real certificate via Certbot):
1.  It is recommended to run Certbot *on the host machine* and modify the `nginx/nginx.conf` to point to the let's encrypt certificates, or use a separate "Nginx Proxy Manager" container.
2.  Alternatively, you can just use the manual Nginx setup described in `DEPLOYMENT_LINUX.md` if you prefer managing Nginx directly on the host rather than in Docker.
