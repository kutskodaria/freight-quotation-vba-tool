CREATE DATABASE IF NOT EXISTS china_belarus_logistics;
USE china_belarus_logistics;

CREATE TABLE IF NOT EXISTS freight_tariffs (
    tariff_id INT AUTO_INCREMENT PRIMARY KEY,
    route_name VARCHAR(100) NOT NULL,
    container_type VARCHAR(20) NOT NULL,
    base_cost_usd DECIMAL(10, 2) NOT NULL
);