FROM nginx:stable-alpine
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY index.html pilot.css pilot.js /usr/share/nginx/html/
RUN chmod 644 /etc/nginx/conf.d/default.conf /usr/share/nginx/html/index.html /usr/share/nginx/html/pilot.css /usr/share/nginx/html/pilot.js
