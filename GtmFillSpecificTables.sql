DECLARE
    v_dobj_id VARCHAR2(100);
    v_table_name VARCHAR2(100);
    v_column_name VARCHAR2(100);
    v_next_value VARCHAR2(100);
    v_count INT;
    v_arg1 NUMBER;
    v_arg2 NUMBER;
    v_arg3 NUMBER;
    v_arg4 NUMBER;
    v_arg5 NUMBER;
    v_arg6 NUMBER;
    v_arg7 NUMBER;
    v_arg8 NUMBER;
    v_ind_value NUMBER;
    
    PROCEDURE select_ind_value(b_name VARCHAR2, ind_id NUMBER, org_name VARCHAR2, field_name VARCHAR2, v_arg OUT NUMBER) IS
    BEGIN
        SELECT DISTINCT IND_VALUE
        INTO v_arg
        FROM RNKIN_READER.DATA_X5
        WHERE BLOK_NAME = b_name
          AND BLOCK_FILLING = 'Стандартно'
          AND IND_ID = ind_id
          AND INSTR(DT, 'год') > 0
          AND CALC_TYPE = 'С учётом предыдущих'
          AND ORG_NAME = org_name
          AND FIELD_NAME = field_name;
    EXCEPTION
        WHEN NO_DATA_FOUND THEN
            v_arg := 0;
    END;
BEGIN
        FOR rec IN (SELECT DISTINCT
                        gtm.PRISMA_ID,
                        rnkin.IND_VALUE,  -- Значение индекса
                        rnkin.ORG_NAME,
                        rnkin.FIELD_NAME,
                        CAST(TRIM(TO_NUMBER(REPLACE(rnkin.DT, 'год', ''))) AS NUMBER) AS YEAR,  -- Преобразование текстового поля с годом в числовое
                        admin.DOBJ_ID  -- Идентификатор связанного объекта
                    FROM 
                        RNKIN_READER.DATA_X5 rnkin
                    JOIN PRS_GTM_DATA gtm
                        ON rnkin.BLOK_NAME = gtm.BLOK_NAME 
                        AND rnkin.BLOCK_FILLING = gtm.BLOCK_FILLING 
                        and rnkin.IND_ID = gtm.IND_ID  -- Соединение двух таблиц по нескольким ключам
                    JOIN PRS_KIN_ADMIN_OBJECTS admin
                        ON rnkin.BLOK_NAME = admin.BLOK_NAME 
                        AND rnkin.BLOCK_FILLING = admin.BLOCK_FILLING 
                        and rnkin.IND_ID = admin.IND_ID
                        and rnkin.FIELD_NAME = admin.FIELD_NAME
                        and rnkin.ORG_NAME = admin.ORG_NAME -- Соединение двух таблиц по нескольким ключам
                    WHERE 
                        rnkin.BLOK_NAME IN (
                            'ЗБС', 'Группировка данных в свод', 'Базовая добыча', 
                            'Ввод новых скважин', 'Ввод прочих новых',
                            'ГРП', 'Перевод и приобщение', 'Вывод скважин из бездействия')  -- Фильтрация по названиям блока
                        AND rnkin.BLOCK_FILLING IN ('Стандартно', 'Подробно')  -- Фильтрация по типу заполнения блока
                        AND INSTR(rnkin.DT, 'год') > 0  -- Проверка наличия текста 'год' в поле dt
                        AND rnkin.CALC_TYPE = 'С учётом предыдущих'  -- Дополнительное условие по типу расчёта
                        and admin.DOBJ_ID = :p_dobj_id  -- Фильтрация по идентификатору объекта, полученному из внешнего источника
                        AND (:p_year IS NULL OR CAST(TRIM(REPLACE(rnkin.DT, 'год', '')) AS NUMBER) = :p_year) -- Сравнение года из dt с заданным годом
                    ) LOOP
            
            CASE 
                WHEN rec.PRISMA_ID LIKE '%ALL%' THEN 
                    v_table_name := 'PRS_DATA_OBJECT_ALL';
                    v_dobj_id := 'DOBJ_ALL_ID';
                    v_next_value := 'PRS_DATA_OBJECT_ALL_ID_S.NEXTVAL';
                    
                WHEN rec.PRISMA_ID LIKE '%BAS%' THEN 
                    v_table_name := 'PRS_DATA_OBJECT_BASE';
                    v_dobj_id := 'DOBJ_BASE_ID';
                    v_next_value := 'PRS_DATA_OBJECT_BASE_ID_S.NEXTVAL';
                    
                WHEN rec.PRISMA_ID LIKE '%VNS%' THEN 
                    v_table_name := 'PRS_DATA_OBJECT_NEW_WELL';
                    v_dobj_id := 'DOBJ_NWELL_ID';
                    v_next_value := 'PRS_DATA_OBJECT_NEW_WELL_ID_S.NEXTVAL';
                    
                WHEN rec.PRISMA_ID LIKE '%ZBS%' THEN 
                    v_table_name := 'PRS_DATA_OBJECT_SWELL';
                    v_dobj_id := 'DOBJ_SWELL_ID';
                    v_next_value := 'PRS_DATA_OBJECT_SWELL_ID_S.NEXTVAL';
                    
                WHEN rec.PRISMA_ID LIKE '%PRO%' THEN 
                    v_table_name := 'PRS_DATA_OBJECT_GTM';
                    v_dobj_id := 'DOBJ_GTM_ID';
                    v_next_value := 'PRS_DATA_OBJECT_GTM_ID_S.NEXTVAL';
                ELSE
                    IF rec.PRISMA_ID = 'NEW_WELL11' THEN
                      v_table_name := 'PRS_DATA_OBJECT_NEW_WELL';
                      v_dobj_id := 'DOBJ_NWELL_ID';
                      v_next_value := 'PRS_DATA_OBJECT_NEW_WELL_ID_S.NEXTVAL';
                      v_column_name := rec.PRISMA_ID;
                    END IF;
            END CASE;
            
            IF rec.PRISMA_ID != 'NEW_WELL11' THEN
              EXECUTE IMMEDIATE 'SELECT column_name 
                                 FROM all_col_comments
                                 WHERE comments LIKE :1 
                                   AND owner = ''PRISM'' 
                                   AND table_name = :2'
              INTO v_column_name
              USING '%' || rec.PRISMA_ID || '%', v_table_name;
            END IF;

            -- Определяем записаны ли уже значения, чтобы не создавать новую запись     
            EXECUTE IMMEDIATE 'SELECT COUNT(*) FROM ' || v_table_name || 
                  ' WHERE DATA_OBJECT_ID = :1 AND YEAR = :2'
            INTO v_count
            USING rec.DOBJ_ID, rec.YEAR;
            
             CASE 
                WHEN rec.PRISMA_ID = 'P5PROP001' THEN
                BEGIN
                   select_ind_value('Ввод прочих новых', 10, rec.ORG_NAME, rec.FIELD_NAME, v_arg1);
                   select_ind_value('ГРП', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
                   select_ind_value('Перевод и приобщение', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg3);
                   select_ind_value('Вывод скважин из бездействия', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg4);
                   
                   v_ind_value := v_arg1 + v_arg2 + v_arg3 + v_arg4;
                END;
                
                WHEN rec.PRISMA_ID = 'P5PROP003' THEN
                BEGIN
                   select_ind_value('Ввод прочих новых', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg1);
                   select_ind_value('ГРП', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
                   select_ind_value('Перевод и приобщение', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg3);
                   select_ind_value('Вывод скважин из бездействия', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg4);
                   
                   v_ind_value := v_arg1 + v_arg2 + v_arg3 + v_arg4;
                END;
                
                WHEN rec.PRISMA_ID = 'P5PROP002' THEN
                BEGIN
                    v_arg1 := 0;
                    select_ind_value('ГРП', 30, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
                    select_ind_value('Перевод и приобщение', 30, rec.ORG_NAME, rec.FIELD_NAME, v_arg3);
                    select_ind_value('Вывод скважин из бездействия', 30, rec.ORG_NAME, rec.FIELD_NAME, v_arg4);  
                   
                    v_ind_value := v_arg1 + v_arg2 + v_arg3 + v_arg4;
                END;
                
                WHEN rec.PRISMA_ID = 'P5PROP029' THEN
                BEGIN
                   select_ind_value('Ввод новых скважин', 310, rec.ORG_NAME, rec.FIELD_NAME, v_arg1);
                   select_ind_value('Ввод прочих новых', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
                   select_ind_value('ГРП', 310, rec.ORG_NAME, rec.FIELD_NAME, v_arg3);
                   select_ind_value('ГРП', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg4);
                   select_ind_value('Перевод и приобщение', 310, rec.ORG_NAME, rec.FIELD_NAME, v_arg1);
                   select_ind_value('Перевод и приобщение', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
                   select_ind_value('Вывод скважин из бездействия', 310, rec.ORG_NAME, rec.FIELD_NAME, v_arg3);
                   select_ind_value('Вывод скважин из бездействия', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg4);
                    
                   v_ind_value := (v_arg1 * v_arg2) + (v_arg3 * v_arg4) + (v_arg5 * v_arg6) + (v_arg7 * v_arg8);
                   v_ind_value := GREATEST(v_ind_value, 0);
                END;
                
                WHEN rec.PRISMA_ID = 'P5PROP019' THEN 
                BEGIN
                   select_ind_value('Ввод прочих новых', 210, rec.ORG_NAME, rec.FIELD_NAME, v_arg1);
                   select_ind_value('Ввод прочих новых', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
                   select_ind_value('ГРП', 210, rec.ORG_NAME, rec.FIELD_NAME, v_arg3);
                   select_ind_value('ГРП', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg4);
                   select_ind_value('Перевод и приобщение', 210, rec.ORG_NAME, rec.FIELD_NAME, v_arg5);
                   select_ind_value('Перевод и приобщение', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg6);
                   select_ind_value('Вывод скважин из бездействия', 210, rec.ORG_NAME, rec.FIELD_NAME, v_arg7);
                   select_ind_value('Вывод скважин из бездействия', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg8);
                    
                   v_ind_value := (v_arg1 * v_arg2) + (v_arg3 * v_arg4) + (v_arg5 * v_arg6) + (v_arg7 * v_arg8);
                   v_ind_value := GREATEST(v_ind_value, 0);
                END;
                
                WHEN rec.PRISMA_ID = 'P5PROP021' THEN 
                  BEGIN
                     select_ind_value('Ввод прочих новых', 211, rec.ORG_NAME, rec.FIELD_NAME, v_arg1);
                     select_ind_value('Ввод прочих новых', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
                     select_ind_value('ГРП', 211, rec.ORG_NAME, rec.FIELD_NAME, v_arg3);
                     select_ind_value('ГРП', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg4);
                     select_ind_value('Перевод и приобщение', 211, rec.ORG_NAME, rec.FIELD_NAME, v_arg5);
                     select_ind_value('Перевод и приобщение', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg6);
                     select_ind_value('Вывод скважин из бездействия', 211, rec.ORG_NAME, rec.FIELD_NAME, v_arg7);
                     select_ind_value('Вывод скважин из бездействия', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg8);
                      
                     v_ind_value := (v_arg1 * v_arg2) + (v_arg3 * v_arg4) + (v_arg5 * v_arg6) + (v_arg7 * v_arg8);
                     v_ind_value := GREATEST(v_ind_value, 0);
                  END;
                
                WHEN rec.PRISMA_ID = 'P5PROP023' THEN
                  BEGIN
                     select_ind_value('Ввод прочих новых', 250, rec.ORG_NAME, rec.FIELD_NAME, v_arg1);
                     select_ind_value('Ввод прочих новых', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
                     select_ind_value('ГРП', 250, rec.ORG_NAME, rec.FIELD_NAME, v_arg3);
                     select_ind_value('ГРП', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg4);
                     select_ind_value('Перевод и приобщение', 250, rec.ORG_NAME, rec.FIELD_NAME, v_arg5);
                     select_ind_value('Перевод и приобщение', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg6);
                     select_ind_value('Вывод скважин из бездействия', 250, rec.ORG_NAME, rec.FIELD_NAME, v_arg7);
                     select_ind_value('Вывод скважин из бездействия', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg8);
                      
                     v_ind_value := (v_arg1 * v_arg2) + (v_arg3 * v_arg4) + (v_arg5 * v_arg6) + (v_arg7 * v_arg8);
                     v_ind_value := GREATEST(v_ind_value, 0);
                  END;
                   
                WHEN rec.PRISMA_ID = 'P5PROP025' THEN 
                  BEGIN
                     select_ind_value('Ввод прочих новых', 251, rec.ORG_NAME, rec.FIELD_NAME, v_arg1);
                     select_ind_value('Ввод прочих новых', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
                     select_ind_value('ГРП', 251, rec.ORG_NAME, rec.FIELD_NAME, v_arg3);
                     select_ind_value('ГРП', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg4);
                     select_ind_value('Перевод и приобщение', 251, rec.ORG_NAME, rec.FIELD_NAME, v_arg5);
                     select_ind_value('Перевод и приобщение', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg6);
                     select_ind_value('Вывод скважин из бездействия', 251, rec.ORG_NAME, rec.FIELD_NAME, v_arg7);
                     select_ind_value('Вывод скважин из бездействия', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg8);
                    
                     v_ind_value := (v_arg1 * v_arg2) + (v_arg3 * v_arg4) + (v_arg5 * v_arg6) + (v_arg7 * v_arg8);
                     v_ind_value := GREATEST(v_ind_value, 0);
                  END;
                
                WHEN rec.PRISMA_ID = 'P5PROP031' THEN 
                  BEGIN
                     select_ind_value('Ввод прочих новых', 910, rec.ORG_NAME, rec.FIELD_NAME, v_arg1);
                     select_ind_value('Ввод прочих новых', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
                     select_ind_value('ГРП', 910, rec.ORG_NAME, rec.FIELD_NAME, v_arg3);
                     select_ind_value('ГРП', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg4);
                     select_ind_value('Перевод и приобщение', 910, rec.ORG_NAME, rec.FIELD_NAME, v_arg5);
                     select_ind_value('Перевод и приобщение', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg6);
                     select_ind_value('Вывод скважин из бездействия', 910, rec.ORG_NAME, rec.FIELD_NAME, v_arg7);
                     select_ind_value('Вывод скважин из бездействия', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg8);
                    
                     v_ind_value := (v_arg1 * v_arg2) + (v_arg3 * v_arg4) + (v_arg5 * v_arg6) + (v_arg7 * v_arg8);
                     v_ind_value := GREATEST(v_ind_value, 0);
                  END;
                
                WHEN rec.PRISMA_ID = 'P5VNS005' THEN 
                  BEGIN
                     select_ind_value('Ввод новых скважин', 20, rec.ORG_NAME, rec.FIELD_NAME, v_arg1);
                     select_ind_value('Ввод новых скважин', 25, rec.ORG_NAME, rec.FIELD_NAME, v_arg2);
  
                     v_ind_value := v_arg1 - v_arg2;
                  END;
                
                ELSE
                  v_ind_value := rec.IND_VALUE; 
              END CASE;
              
            IF v_count = 0 THEN
              EXECUTE IMMEDIATE 'INSERT INTO ' || v_table_name || ' ('
                      || v_dobj_id || ', 
                      DATA_OBJECT_ID, 
                      YEAR, 
                      ' || v_column_name || '
                  ) VALUES(
                      ' || v_next_value || ',
                      :1,
                      :2,
                      :3)' USING rec.DOBJ_ID, rec.YEAR, v_ind_value;
                      
              INSERT INTO PRS_LOG (LOG_ID, LOG_TYPE_ID, LOG_USR_ID, OBJECT_ID, PARENT_OBJECT_ID, LOG_DETAILS)
              VALUES (PRS_LOG_ID_S.NEXTVAL, 40001, :p_usrid, rec.DOBJ_ID, NULL, NULL);
              
            ELSE
              EXECUTE IMMEDIATE 'UPDATE ' || v_table_name || ' 
                  SET ' || v_column_name || ' = :1
                  WHERE DATA_OBJECT_ID = :2 and YEAR = :3' USING rec.IND_VALUE, rec.DOBJ_ID, rec.YEAR;
            END IF;
        END LOOP;

    -- Если все команды выполнены успешно, зафиксируем изменения
    COMMIT;

EXCEPTION
    WHEN OTHERS THEN
        -- Если произошла ошибка, откатим все изменения
        ROLLBACK;
        
        RAISE_APPLICATION_ERROR(-20001, SQLERRM);
END;