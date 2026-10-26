FROM nginx:stable-alpine
COPY oracle-original-canary/nginx.conf /etc/nginx/conf.d/default.conf
COPY . /usr/share/nginx/html/
RUN find /usr/share/nginx/html -type d -exec chmod 755 {} + \
  && find /usr/share/nginx/html -type f -exec chmod 644 {} + \
  && chmod 644 /etc/nginx/conf.d/default.conf
