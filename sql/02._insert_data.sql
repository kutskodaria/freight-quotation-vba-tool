USE china_belarus_logistics;

TRUNCATE TABLE freight_tariffs;

INSERT INTO freight_tariffs (route_name, container_type, base_cost_usd) VALUES 
('Ningbo - Kolyadichi', '40ft HC', 3200.00),
('Ningbo - Kolyadichi', '20ft GP', 2400.00),
('Shanghai - Kolyadichi', '40ft HC', 3400.00),
('Shanghai - Kolyadichi', '20ft GP', 2600.00),
('Guangzhou - Kolyadichi', '40ft HC', 3600.00),
('Guangzhou - Kolyadichi', '20ft GP', 2800.00);

SELECT * FROM freight_tariffs;