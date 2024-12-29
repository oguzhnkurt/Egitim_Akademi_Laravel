# Dockerfile
FROM php:8.1-fpm

# Gerekli uzantıları yükleyin
RUN apt-get update && apt-get install -y \
    libpng-dev \
    libjpeg-dev \
    libfreetype6-dev \
    libzip-dev \
    unzip \
    && docker-php-ext-configure gd --with-freetype --with-jpeg \
    && docker-php-ext-install gd zip

# Composer'ı yükleyin
COPY --from=composer:latest /usr/bin/composer /usr/bin/composer

# Çalışma dizinini ayarlayın
WORKDIR /var/www/html

# Uygulama dosyalarını kopyalayın
COPY . .

# Bağımlılıkları yükleyin
RUN composer install

# Laravel için gerekli izinleri ayarlayın
RUN chown -R www-data:www-data /var/www/html/storage /var/www/html/bootstrap/cache