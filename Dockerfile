# syntax=docker/dockerfile:1.6

FROM python:3.12-slim AS build
WORKDIR /src
RUN pip install --no-cache-dir "zensical==0.0.29"
COPY zensical.toml ./
COPY docs ./docs
RUN zensical build

FROM nginx:1.27-alpine AS runtime
COPY --from=build /src/site /usr/share/nginx/html
EXPOSE 80
