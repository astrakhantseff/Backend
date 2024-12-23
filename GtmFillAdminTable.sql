DECLARE
    -- Переменная для хранения счета
    v_count NUMBER;
    v_dobj_id NUMBER(18,0);

    TYPE t_dobj_id IS TABLE OF PRS_DATA_OBJECT.DOBJ_ID%TYPE;
    v_dobj_ids t_dobj_id;

    CURSOR record_cursor IS
      SELECT DISTINCT 
          rnkin.BLOK_NAME,
          rnkin.BLOCK_FILLING,
          gtm.BLOCK_ID, 
          gtm.BLOCK_FILLING_ID, 
          rnkin.IND_ID,
          rnkin.ORG_NAME,
          rnkin.FIELD_NAME,
          MIN(CAST(REPLACE(rnkin.DT, 'год', '') AS NUMBER)) AS FIRST_YEAR
      FROM 
          RNKIN_READER.DATA_X5 rnkin
      JOIN PRS_GTM_DATA gtm ON 
          rnkin.BLOK_NAME = gtm.BLOK_NAME 
          AND rnkin.BLOCK_FILLING = gtm.BLOCK_FILLING 
          AND rnkin.IND_ID = gtm.IND_ID
      where 
          rnkin.BLOK_NAME IN ('ЗБС', 'Группировка данных в свод', 'Базовая добыча', 'Ввод новых скважин', 'Ввод прочих новых',
                              'ГРП', 'Перевод и приобщение', 'Вывод скважин из бездействия')
          AND rnkin.BLOCK_FILLING IN ('Стандартно', 'Подробно')
          AND INSTR(rnkin.DT, 'год') > 0
          AND rnkin.CALC_TYPE = 'С учётом предыдущих'
      GROUP BY 
          rnkin.BLOK_NAME,
          rnkin.BLOCK_FILLING,
          gtm.BLOCK_ID, 
          gtm.BLOCK_FILLING_ID, 
          rnkin.IND_ID,
          rnkin.ORG_NAME,
          rnkin.FIELD_NAME;
BEGIN
    
    -- Выполним запрос и сразу подсчитаем количество строк, которые не совпадают
    SELECT COUNT(*)
    INTO v_count -- Сохраняем результат в переменную
    FROM (
        -- Основной запрос, который использует MINUS для поиска несовпадающих строк
        SELECT BLOK_NAME, BLOCK_FILLING, IND_ID
        FROM PRS_KIN_ADMIN_OBJECTS
        MINUS
        SELECT DISTINCT 
            rnkin.BLOK_NAME,
            rnkin.BLOCK_FILLING,
            rnkin.IND_ID
        FROM
            RNKIN_READER.DATA_X5 rnkin
        JOIN PRS_KIN_ADMIN_OBJECTS admin ON
            rnkin.BLOK_NAME = admin.BLOK_NAME 
            AND rnkin.BLOCK_FILLING = admin.BLOCK_FILLING 
            AND rnkin.IND_ID = admin.IND_ID
            AND rnkin.ORG_NAME = admin.ORG_NAME
            AND rnkin.FIELD_NAME = admin.FIELD_NAME
        where 
            rnkin.BLOK_NAME IN ('ЗБС', 'Группировка данных в свод', 'Базовая добыча', 'Ввод новых скважин', 'Ввод прочих новых',
                                'ГРП', 'Перевод и приобщение', 'Вывод скважин из бездействия')
            AND rnkin.BLOCK_FILLING IN ('Стандартно', 'Подробно')
            AND INSTR(rnkin.DT, 'год') > 0
            AND rnkin.CALC_TYPE = 'С учётом предыдущих'
        GROUP BY 
            rnkin.BLOK_NAME,
            rnkin.BLOCK_FILLING,
            rnkin.IND_ID
    );
    
    IF v_count > 0 THEN
      -- Сбор всех DOBJ_ID из PRS_KIN_ADMIN_OBJECTS
      SELECT DOBJ_ID BULK COLLECT INTO v_dobj_ids FROM PRS_KIN_ADMIN_OBJECTS;
      
      -- Удаление связанных данных из всех соответствующих таблиц, где DATA_OBJECT_ID совпадает
      FORALL i IN v_dobj_ids.FIRST .. v_dobj_ids.LAST
          DELETE FROM PRS_DATA_OBJECT_ALL WHERE DATA_OBJECT_ID = v_dobj_ids(i);
    
      FORALL i IN v_dobj_ids.FIRST .. v_dobj_ids.LAST
          DELETE FROM PRS_DATA_OBJECT_BASE WHERE DATA_OBJECT_ID = v_dobj_ids(i);

      FORALL i IN v_dobj_ids.FIRST .. v_dobj_ids.LAST
          DELETE FROM PRS_DATA_OBJECT_NEW_WELL WHERE DATA_OBJECT_ID = v_dobj_ids(i);

      FORALL i IN v_dobj_ids.FIRST .. v_dobj_ids.LAST
          DELETE FROM PRS_DATA_OBJECT_SWELL WHERE DATA_OBJECT_ID = v_dobj_ids(i);

      FORALL i IN v_dobj_ids.FIRST .. v_dobj_ids.LAST
          DELETE FROM PRS_DATA_OBJECT_GTM WHERE DATA_OBJECT_ID = v_dobj_ids(i);

      FORALL i IN v_dobj_ids.FIRST .. v_dobj_ids.LAST
          DELETE FROM PRS_DATA_OBJECT_SRC WHERE DATA_OBJECT_CHILD = v_dobj_ids(i);

      FORALL i IN v_dobj_ids.FIRST .. v_dobj_ids.LAST
          DELETE FROM PRS_DATA_OBJECT WHERE DOBJ_ID = v_dobj_ids(i);
          
      EXECUTE IMMEDIATE 'TRUNCATE TABLE PRS_KIN_ADMIN_OBJECTS';

    END IF;
    
    FOR rec IN record_cursor LOOP
      
      BEGIN
          -- Пытаемся выбрать DOBJ_ID из таблицы
          SELECT DOBJ_ID 
          INTO v_dobj_id
          FROM PRS_KIN_ADMIN_OBJECTS
          WHERE BLOK_NAME = rec.BLOK_NAME
            AND BLOCK_FILLING = rec.BLOCK_FILLING
            AND IND_ID = rec.IND_ID
            AND FIELD_NAME = rec.FIELD_NAME
            AND ORG_NAME = rec.ORG_NAME;
      
      EXCEPTION
          -- Отлавливаем ситуацию, когда ни одна строка не была найдена
          WHEN NO_DATA_FOUND THEN
              v_dobj_id := NULL;
      END;
      
      IF v_dobj_id is null THEN
        INSERT INTO PRS_DATA_OBJECT (
                  DOBJ_ID,
                  DATAOBJ_STATE_ID,
                  TECH_ERROR_STATE,
                  IS_APPROVAL,
                  SMN_ERROR_STATE,
                  SRC_ORGANIZATION, 
                  SRC_FIELD, 
                  SRC_PROCESS, 
                  SRC_VARIANT, 
                  SRC_VERSION,
                  SRC_DATA_FORMAT,
                  FIRST_YEAR
              ) VALUES (
                  PRS_DATA_OBJECT_ID_S.NEXTVAL, 
                  1,
                  1,
                  3,
                  1,
                  rec.ORG_NAME, 
                  rec.FIELD_NAME, 
                  'БП', 
                  'СД', 
                  TO_DATE(TO_CHAR(SYSDATE, 'DDMMYYYY')),
                  'X5',
                  rec.FIRST_YEAR
              );
          
          -- Вставка связанной записи в таблицу PRS_DATA_OBJECT_SRC
          INSERT INTO PRS_DATA_OBJECT_SRC (
                  DOBJ_SRC_ID,
                  DATA_OBJECT_CHILD
              ) VALUES (
                  PRS_DATA_OBJECT_SRC_ID_S.NEXTVAL,
                  PRS_DATA_OBJECT_ID_S.CURRVAL
              );
              
          INSERT INTO PRS_KIN_ADMIN_OBJECTS(
              BLOK_NAME,
              BLOCK_FILLING,
              IND_ID,
              DOBJ_ID,
              ORG_NAME,
              FIELD_NAME
          ) VALUES (
              rec.BLOK_NAME,
              rec.BLOCK_FILLING,
              rec.IND_ID,
              PRS_DATA_OBJECT_ID_S.CURRVAL,
              rec.ORG_NAME,
              rec.FIELD_NAME
          );
          
          INSERT INTO PRS_LOG (LOG_ID, LOG_TYPE_ID, LOG_USR_ID, OBJECT_ID, PARENT_OBJECT_ID, LOG_DETAILS) 
          VALUES (PRS_LOG_ID_S.NEXTVAL, 40001, 1012, PRS_DATA_OBJECT_ID_S.CURRVAL, NULL, NULL);

        ELSE
          -- Обновление записи в таблице PRS_DATA_OBJECT, если запись существует
          UPDATE PRS_DATA_OBJECT
          SET 
            DATAOBJ_STATE_ID = 1,
            TECH_ERROR_STATE = 1,
            IS_APPROVAL = 3,
            SMN_ERROR_STATE = 1,
            SRC_ORGANIZATION = rec.ORG_NAME, 
            SRC_FIELD = rec.FIELD_NAME, 
            SRC_PROCESS = 'БП', 
            SRC_VARIANT = 'СД', 
            SRC_VERSION = TO_DATE(TO_CHAR(SYSDATE, 'DDMMYYYY')), 
            SRC_DATA_FORMAT = 'X5',
            FIRST_YEAR = rec.FIRST_YEAR
          WHERE 
            DOBJ_ID = v_dobj_id;

          --Обновление записи в таблице PRS_GTM_DATA
          UPDATE PRS_KIN_ADMIN_OBJECTS
          SET 
            ORG_NAME = rec.ORG_NAME,
            FIELD_NAME = rec.FIELD_NAME
          WHERE 
            DOBJ_ID = v_dobj_id;

      END IF;
    END LOOP;

    COMMIT;
    
EXCEPTION
    WHEN OTHERS THEN
        -- Если произошла ошибка, откатим все изменения
        ROLLBACK;
        
        RAISE_APPLICATION_ERROR(-20001, SQLERRM);
END;