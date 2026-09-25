-- ==============================================================================
-- Script de Limpeza e Correção de Localidades com 'nan' e 'Não encontrado no AD'
-- Banco: chamados.db
-- ==============================================================================

-- 1. Remove registros espúrios com ID 'nan' ou nulo
DELETE FROM comentarios WHERE chamado_id = 'nan' OR chamado_id IS NULL;
DELETE FROM chamados WHERE id = 'nan' OR id IS NULL;

-- 2. Limpa valores 'nan' ou literais nulos nas colunas de localização
UPDATE chamados 
SET cidade_predio = '' 
WHERE LOWER(TRIM(cidade_predio)) IN ('nan', 'none', 'null', '<na>') OR cidade_predio IS NULL;

UPDATE chamados 
SET unidade = '' 
WHERE LOWER(TRIM(unidade)) IN ('nan', 'none', 'null', '<na>') 
   OR LOWER(unidade) LIKE '%não encontrad% no ad%' 
   OR LOWER(unidade) LIKE '%nao encontrad% no ad%'
   OR unidade IS NULL;

-- 3. Atualiza localidade_fisica que contenha 'nan' ou referências a erro do AD
UPDATE chamados 
SET localidade_fisica = 'Não identificada'
WHERE LOWER(localidade_fisica) LIKE '%nan%' 
   OR LOWER(localidade_fisica) LIKE '%não encontrad%' 
   OR LOWER(localidade_fisica) LIKE '%nao encontrad%'
   OR LOWER(TRIM(localidade_fisica)) IN ('none', 'null', '<na>', '', 'n/d')
   OR localidade_fisica IS NULL;

-- 4. Para chamados onde cidade_predio está preenchido,
-- define a localidade_fisica como a Cidade / Prédio (removendo sufixos desnecessários como ' - Sede')
UPDATE chamados
SET localidade_fisica = TRIM(REPLACE(cidade_predio, ' - Sede', ''))
WHERE cidade_predio IS NOT NULL AND TRIM(cidade_predio) != ''
  AND (
    localidade_fisica = 'Não identificada'
    OR localidade_fisica LIKE '% - %ª PJ%'
    OR localidade_fisica LIKE '% - %º PJ%'
    OR localidade_fisica LIKE '% - Promotoria%'
    OR localidade_fisica LIKE '% - Procuradoria%'
  );

-- 5. Padroniza localidades de Campo Grande para não concatenar setores internos na localidade física
UPDATE chamados
SET localidade_fisica = 'Campo Grande - PGJ'
WHERE localidade_fisica LIKE 'Campo Grande - PGJ - %';

UPDATE chamados
SET localidade_fisica = 'Campo Grande - DMP'
WHERE localidade_fisica LIKE 'Campo Grande - DMP - %';

UPDATE chamados
SET localidade_fisica = 'Campo Grande - Rua da Paz'
WHERE localidade_fisica LIKE 'Campo Grande - Rua da Paz - %';

UPDATE chamados
SET localidade_fisica = 'Campo Grande - Chácara Cachoeira'
WHERE localidade_fisica LIKE 'Campo Grande - Chácara Cachoeira - %';
