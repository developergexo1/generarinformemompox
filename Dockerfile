# Check out https://hub.docker.com/_/node to select a new base image
FROM node:22.13.0

# Create app directory (with user `node`)
RUN mkdir -p /home/node/app

WORKDIR /home/node/app

# Bind to all network interfaces so that it can be mapped to the host OS
ENV HOST=0.0.0.0 PORT=1500

EXPOSE ${PORT}

RUN apt update

# Install nano
RUN apt install nano

# Bundle app source code
COPY . .

ENV TZ=America/Bogota

RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone

RUN npm install

CMD ["npm", "start"]