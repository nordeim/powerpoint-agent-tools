### Plan (goals and decisions)

- Goal: provide a compact, developer-friendly docker-compose.yml that:
  - builds the image from your Dockerfile,
  - starts the container detached and keeps it alive,
  - mounts the repository into /app for live editing,
  - exposes common ports (configurable) for services you’ll run during development,
  - makes it trivial to open an interactive shell using docker compose exec.
- Decisions:
  - Use Compose v3.8 for broad compatibility.
  - Use a simple long-running command (tail -f /dev/null) as the default for dev so container stays up.
  - Mount the repo as a bind mount with :delegated for better performance on macOS/Linux.
  - Provide environment passthrough for PYTHONPATH and an optional DEV_PORT variable.
  - Keep the service runnable as non-root by optionally using service user mapping later (comment included).

---

### docker-compose.yml (dev-friendly)

```yaml
version: "3.8"

services:
  ppt-agent-tools:
    build:
      context: .
      dockerfile: docker/Dockerfile
      args:
        # allow build-time override if you use ARG PYTHONPATH in Dockerfile
        PYTHONPATH: ""
    image: jeremiahwong/ppt-agent-tools:dev
    container_name: ppt-agent-tools-dev
    command: ["tail", "-f", "/dev/null"]
    working_dir: /app
    environment:
      - PYTHONPATH=/app
      - DEV_PORT=8000
    ports:
      - "8000:8000"    # change or remove as needed
    volumes:
      - ".:/app:delegated"
    restart: unless-stopped
    tty: true
```

---

### How to use (exact commands)

- Start detached (builds image if needed):
  - docker compose up -d
- Rebuild and recreate (useful after Dockerfile changes):
  - docker compose up -d --build
- Rebuild without cache (ensure fresh layers):
  - docker compose build --no-cache
  - docker compose up -d
- Open an interactive shell inside the running container:
  - docker compose exec ppt-agent-tools /bin/bash
  - If bash missing: docker compose exec ppt-agent-tools /bin/sh
- Stop and remove containers:
  - docker compose down
- View logs:
  - docker compose logs -f ppt-agent-tools

---

### Notes, hardening and recommendations

- If you prefer a semantic tag for development, change image to jeremiahwong/ppt-agent-tools:latest or your chosen tag.
- For non-root development convenience, create a user in the Dockerfile and use the `user: "1000:1000"` compose key or `--user` in run commands to match host UID/GID.
- If your app exposes other ports, update the ports list or set DEV_PORT via environment and use it in your run commands.
- For production or CI, do not mount the repo; build an immutable image and run the actual service command instead of tail.
- Add healthcheck in the compose file when you replace the tail command with a real service to enable orchestrated restarts.

---

### Quick checklist to adopt this in your repo

- [ ] Add the docker-compose.yml to repo root.
- [ ] Ensure top-level .dockerignore does not exclude files you need inside the container (requirements.txt, bin/uv).
- [ ] Run docker compose up -d and verify container name appears in docker ps.
- [ ] Exec into container and confirm /app maps to repo and PYTHONPATH contains /app.

If you want, I can produce a variant that adds a named volume for persistent data, a healthcheck, or an example service entry (e.g., uvicorn command) wired to DEV_PORT. Which variant do you prefer?

---

https://copilot.microsoft.com/shares/FCv2y7B4t6yznc3smyiR1  
https://copilot.microsoft.com/shares/TkPuwwdbopyanEuz2rD4t  
