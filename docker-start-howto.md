### Why the container exited immediately

Your image’s default command is a shell that runs and then exits because there is no long‑running process keeping the container alive. The image metadata shows CMD ["/bin/bash"] so when Docker starts the container it runs /bin/bash non‑interactively, the shell exits immediately, and the container stops with exit code 0.

Quick diagnostics
- See the container exit code and last logs:  
  docker ps -a --filter id=a93ff603efc8 --format "table {{.ID}}\t{{.Image}}\t{{.Status}}" && docker logs a93ff603efc8
- Inspect configured command and entrypoint:  
  docker image inspect jeremiahwong/ppt-agent-tools:latest --format '{{json .Config.Cmd}} {{json .Config.Entrypoint}}'

---

### How to start a container that stays running

Choose one of these approaches depending on whether you want the container to run a real service or just stay alive for interactive access.

- Run a long‑lived command (recommended for quick dev sessions):  
  docker run -d --name ppt-agent-tools \
    --restart unless-stopped \
    jeremiahwong/ppt-agent-tools:latest \
    tail -f /dev/null

- Start with bash in the background (keeps it alive):  
  docker run -d --name ppt-agent-tools \
    jeremiahwong/ppt-agent-tools:latest \
    /bin/bash -c "while true; do sleep 86400; done"

- Update the Dockerfile to provide a dev-friendly default (persists across runs):  
  Replace CMD ["/bin/bash"] with CMD ["tail","-f","/dev/null"] or use your real service command as CMD.

---

### How to open a shell anytime

- If the container is already running, use docker exec for an interactive shell:  
  docker exec -it ppt-agent-tools /bin/bash  
  If bash isn’t present, use /bin/sh.

- Start a new ephemeral interactive container (no background):  
  docker run --rm -it -v "$PWD":/app:delegated --entrypoint /bin/bash jeremiahwong/ppt-agent-tools:latest

- Start container detached then attach shell later:  
  docker run -d --name ppt-agent-tools jeremiahwong/ppt-agent-tools:latest tail -f /dev/null  
  docker exec -it ppt-agent-tools /bin/bash

---

### Useful lifecycle and debug commands

- See running containers: docker ps  
- See all containers and their exit reasons: docker ps -a  
- Inspect exit code: docker inspect --format='{{.State.ExitCode}} {{.State.Status}}' <container-id>  
- View last logs: docker logs <container-id>  
- Start/stop/remove: docker start <name> ; docker stop <name> ; docker rm <name>

---

### Recommended development workflow

1. For development mount your repo and keep container alive:
   docker run -d --name ppt-agent-tools -v "$PWD":/app:delegated -w /app jeremiahwong/ppt-agent-tools:latest tail -f /dev/null

2. Exec into container to test and iterate:
   docker exec -it ppt-agent-tools /bin/bash

3. When you have a real service, set that as CMD in the Dockerfile and use proper process management (tini) and a non‑root user before promoting to production.

---

### Quick checklist to fix your current run

- [ ] Start container with a long‑running command (tail -f /dev/null) or change Dockerfile CMD.  
- [ ] Use docker exec -it <name> /bin/bash to get shell access.  
- [ ] Inspect logs and image Config if unexpected exits persist.  

---

https://copilot.microsoft.com/shares/djU8DncibzmtEfM8bhcgY  
https://copilot.microsoft.com/shares/mNXVMwZQwoudMBnwgWH6D  

