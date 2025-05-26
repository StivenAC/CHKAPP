load_query_des = """SELECT
            SIIAPP_Des.N_control
            ,SIIAPP_Des.Fecha_Soli
            ,SIIAPP_Des.N_cotizacion
            ,SIIAPP_Des.PT
            ,SIIAPP_Des.Producto
            ,SIIAPP_Des.Cliente
            ,SIIAPP_Des.Notif_Sanitaria
            ,SIIAPP_Des.Estado_M
            ,SIIAPP_Des.Dis_Des
            ,SIIAPP_Des.Obs
            ,SIIAPP_Des.NIT
            ,SIIAPP_Des.Comercial
            ,SIIAPP_Des.Marca
            FROM dbo.SIIAPP_Des
            """
create_query_des = """
                INSERT INTO dbo.SIIAPP_Des (
                    N_control, 
                    Fecha_Soli, 
                    N_cotizacion, 
                    PT, 
                    Producto, 
                    Cliente,
                    Notif_Sanitaria,
                    Estado_M,
                    Dis_Des,
                    Obs,
                    NIT,
                    Comercial,
                    Marca
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
query_update_des = """
                        UPDATE dbo.SIIAPP_Des
                        SET 
                            N_cotizacion = ?, 
                            PT = ?, 
                            Producto = ?, 
                            Cliente = ?,
                            Notif_Sanitaria = ?,
                            Estado_M = ?, 
                            Dis_Des = ?, 
                            Obs = ?,
                            NIT = ?,
                            Comercial = ?,                     
                            Marca = ?                     
                        WHERE N_control = ?
                    """
load_query_costos = """SELECT
                    SIIAPP_Bod.N_Control
                    ,SIIAPP_Des.Fecha_Soli
                    ,SIIAPP_Des.N_cotizacion
                    ,SIIAPP_Des.PT
                    ,SIIAPP_Des.Producto
                    ,SIIAPP_Bod.Bottle_Cod
                    ,SIIAPP_Bod.Envase
                    ,SIIAPP_Bod.Bottle_Sum
                    ,SIIAPP_Bod.Bottle_Color
                    ,SIIAPP_Bod.Bottle_Material
                    ,SIIAPP_Bod.Bottle_ML
                    ,SIIAPP_Bod.Cap_Cod
                    ,SIIAPP_Bod.Tapa
                    ,SIIAPP_Bod.Cap_Sum
                    ,SIIAPP_Bod.Cap_Color
                    ,SIIAPP_Bod.Cap_Material
                    ,SIIAPP_Bod.Foil_Cod
                    ,SIIAPP_Bod.Foil
                    ,SIIAPP_Bod.Foil_Type
                    ,SIIAPP_Bod.FB_Cod
                    ,SIIAPP_Bod.Funda_Banda
                    ,SIIAPP_Bod.Ubicacion_Termoduc
                    ,SIIAPP_Bod.Etiq_Cod
                    ,SIIAPP_Bod.Etiqueta
                    ,SIIAPP_Bod.Box_Cod
                    ,SIIAPP_Bod.Box_Folding
                    ,SIIAPP_Bod.Box_Sum
                    ,SIIAPP_Bod.Dis_Bod
                    ,SIIAPP_Bod.Bottle_Type
                    ,SIIAPP_Bod.Bottle_Cuello
                    ,SIIAPP_Bod.Bottle_Sellado
                    ,SIIAPP_Bod.Cap_Type
                    ,SIIAPP_Bod.Cap_Valvule
                    ,SIIAPP_Bod.Cap_Overcap
                    ,SIIAPP_Bod.Cap_Style
                    ,SIIAPP_Bod.Cap_Casquete_End
                    ,SIIAPP_Bod.Cap_Casquete_Color
                    ,SIIAPP_Bod.Bottle_Deco
                    ,SIIAPP_Bod.Bottle_Deco_Cod
                    ,SIIAPP_Bod.Bottle_Deco_Desc
                    ,SIIAPP_Bod.Cap_Deco
                    ,SIIAPP_Bod.Cap_Deco_Cod
                    ,SIIAPP_Bod.Cap_Deco_Desc
                    ,SIIAPP_Bod.Cap_Deco_Etiq
                    ,SIIAPP_Bod.Cap_Deco_Etiq_Cod
                    ,SIIAPP_Bod.Foil_Bool
                    ,SIIAPP_Bod.FB_Bool
                    ,SIIAPP_Bod.Escurr_Bool
                    ,SIIAPP_Bod.Escurr_Cod
                    ,SIIAPP_Bod.Escurr_Desc
                    ,SIIAPP_Bod.Capil_Bool
                    ,SIIAPP_Bod.Capil_Obs
                    ,SIIAPP_Bod.Capil_Dimen
                    ,SIIAPP_Bod.Acces_Bool
                    ,SIIAPP_Bod.Acces_Which
                    FROM dbo.SIIAPP_Bod
                    INNER JOIN dbo.SIIAPP_Des
                    ON SIIAPP_Bod.N_Control = SIIAPP_Des.N_control
                    """
update_query_costos = """
                    UPDATE dbo.SIIAPP_Bod SET
                        Bottle_Cod = ?,
                        Envase = ?,
                        Bottle_Sum = ?,
                        Bottle_Color = ?,
                        Bottle_Material = ?,
                        Bottle_ML = ?,
                        Cap_Cod = ?,
                        Tapa = ?,
                        Cap_Sum = ?,
                        Cap_Color = ?,
                        Cap_Material = ?,
                        Foil_Cod = ?,
                        Foil = ?,
                        Foil_Type = ?,
                        FB_Cod = ?,
                        Funda_Banda = ?,
                        Ubicacion_Termoduc = ?,
                        Etiq_Cod = ?,
                        Etiqueta = ?,
                        Box_Cod = ?,
                        Box_Folding = ?,
                        Box_Sum = ?,
                        Dis_Bod = ?,
                        Bottle_Type = ?,
                        Bottle_Cuello = ?,
                        Bottle_Sellado = ?,
                        Cap_Type = ?,
                        Cap_Valvule = ?,
                        Cap_Overcap = ?,
                        Cap_Style = ?,
                        Cap_Casquete_End = ?,
                        Cap_Casquete_Color = ?,
                        Bottle_Deco =?,
                        Bottle_Deco_Cod =?,
                        Bottle_Deco_Desc =?,
                        Cap_Deco =?,
                        Cap_Deco_Cod =?,
                        Cap_Deco_Desc =?,
                        Cap_Deco_Etiq =?,
                        Cap_Deco_Etiq_Cod =?,
                        Foil_Bool =?,
                        FB_Bool =?,
                        Escurr_Bool =?,
                        Escurr_Cod =?,
                        Escurr_Desc =?,
                        Capil_Bool =?,
                        Capil_Obs =?,
                        Capil_Dimen =?,
                        Acces_Bool =?,
                        Acces_Which =?
                    WHERE N_Control = ?;
                    """
load_query_calidad = """SELECT
                 SIIAPP_Cal.N_Control
                ,SIIAPP_Des.Fecha_Soli
                ,SIIAPP_Des.N_cotizacion
                ,SIIAPP_Des.PT
                ,SIIAPP_Des.Producto
                ,SIIAPP_Cal.Mesofilos
                ,SIIAPP_Cal.E_Coli
                ,SIIAPP_Cal.S_Aureus
                ,SIIAPP_Cal.P_Aureoginosa
                ,SIIAPP_Cal.Moho_Levadura
                ,SIIAPP_Cal.Micro_Obs
                ,SIIAPP_Cal.Cumple_Microbiologia
                ,SIIAPP_Cal.Densidad
                ,SIIAPP_Cal.Ph
                ,SIIAPP_Cal.Contenido
                ,SIIAPP_Cal.Color
                ,SIIAPP_Cal.Textura
                ,SIIAPP_Cal.Olor
                ,SIIAPP_Cal.Viscosidad
                ,SIIAPP_Cal.Rpm
                ,SIIAPP_Cal.Aguja
                ,SIIAPP_Cal.Torque
                ,SIIAPP_Cal.Apt_Container
                ,SIIAPP_Cal.Cont_Obs
                ,SIIAPP_Cal.Diligenciado
                ,SIIAPP_Cal.Titular_Nsoc
                ,SIIAPP_Cal.Uso_Exclu
                ,SIIAPP_Cal.Forma_Cosm
                FROM dbo.SIIAPP_Cal
                INNER JOIN dbo.SIIAPP_Des
                ON SIIAPP_Cal.N_Control = SIIAPP_Des.N_control
            """
update_query_calidad = """
                    UPDATE dbo.SIIAPP_Cal SET
                        Mesofilos = ?,
                        E_Coli = ?,
                        S_Aureus = ?,
                        P_Aureoginosa = ?,
                        Moho_Levadura = ?,
                        Micro_Obs = ?,
                        Cumple_Microbiologia = ?,
                        Densidad = ?,
                        Ph = ?,
                        Contenido = ?,
                        Color = ?,
                        Textura = ?,
                        Olor = ?,
                        Viscosidad = ?,
                        Rpm = ?,
                        Aguja = ?,
                        Torque = ?,
                        Apt_Container = ?,
                        Cont_Obs = ?,
                        Diligenciado = ?,
                        Titular_Nsoc = ?,
                        Uso_Exclu = ?,
                        Forma_Cosm = ?
                    WHERE N_Control = ?;
                    """                        
load_query_produccion = """SELECT
            SIIAPP_Prod.N_Control
            ,SIIAPP_Des.Fecha_Soli
            ,SIIAPP_Des.N_cotizacion
            ,SIIAPP_Des.PT
            ,SIIAPP_Des.Producto
            ,SIIAPP_Prod.Cod_Caja
            ,SIIAPP_Prod.Box_Corr
            ,SIIAPP_Prod.Caja_Embalaje
            ,SIIAPP_Prod.FB_Dimensiones_Adec
            ,SIIAPP_Prod.FB_Tunel_Temp
            ,SIIAPP_Prod.FB_Tunel_Speed
            ,SIIAPP_Prod.FB_Test_Tool
            ,SIIAPP_Prod.FB_obs
            ,SIIAPP_Prod.SCE_Temp
            ,SIIAPP_Prod.SCE_Temp_Time
            ,SIIAPP_Prod.SCE_Obs
            ,SIIAPP_Prod.SCE_Hermt
            ,SIIAPP_Prod.Ubi_Lote
            ,SIIAPP_Prod.Etiqueta_Manual
            ,SIIAPP_Prod.Dis_Prod
            ,SIIAPP_Prod.Metod_Fab
            ,SIIAPP_Prod.Diligenciado
            ,SIIAPP_Prod.Box_Units
            ,SIIAPP_Prod.Comb_Dimensions
            ,SIIAPP_Prod.Box_Print_Texture
            ,SIIAPP_Prod.Comb_Bool
            ,SIIAPP_Prod.Rotulo
            ,SIIAPP_Prod.Rotulo_Text
            ,SIIAPP_Prod.Comb_Key
            FROM dbo.SIIAPP_Prod
            INNER JOIN dbo.SIIAPP_Des
            ON SIIAPP_Prod.N_Control = SIIAPP_Des.N_control
            """
edit_query_produccion = """
                        UPDATE dbo.SIIAPP_Prod SET
                        Cod_Caja = ?,
                        Box_Corr = ?,
                        Caja_Embalaje = ?,
                        FB_Dimensiones_Adec = ?,
                        FB_Tunel_Temp = ?,
                        FB_Tunel_Speed = ?,
                        FB_Test_Tool = ?,
                        FB_obs = ?,
                        SCE_Temp = ?,
                        SCE_Temp_Time = ?,
                        SCE_Obs = ?,
                        SCE_Hermt = ?,
                        Ubi_Lote = ?,
                        Etiqueta_Manual = ?,
                        Dis_Prod = ?,
                        Metod_Fab = ?,
                        Diligenciado = ?,
                        Box_Units = ?,
                        Comb_Dimensions = ?,
                        Box_Print_Texture = ?,
                        Comb_Bool = ?,
                        Rotulo = ?,
                        Rotulo_Text = ?,
                        Comb_Key = ?
                    WHERE N_Control = ?;
                    """
pdf_query = """  SELECT
                    SIIAPP_Des.N_control AS [0]
                  ,SIIAPP_Des.Fecha_Soli AS [1]
                  ,SIIAPP_Des.N_cotizacion AS [2]
                  ,SIIAPP_Des.PT AS [3]
                  ,SIIAPP_Des.Producto AS [4]
                  ,SIIAPP_Des.Cliente AS [5]
                  ,SIIAPP_Des.Notif_Sanitaria AS [6]
                  ,SIIAPP_Des.Obs AS [7]
                  ,SIIAPP_Bod.Bottle_Cod AS [8]
                  ,SIIAPP_Bod.Envase AS [9]
                  ,SIIAPP_Bod.Bottle_Sum AS [10]
                  ,SIIAPP_Bod.Bottle_Color AS [11]
                  ,SIIAPP_Bod.Bottle_Material AS [12]
                  ,SIIAPP_Bod.Bottle_ML AS [13]
                  ,SIIAPP_Bod.Cap_Cod AS [14]
                  ,SIIAPP_Bod.Tapa AS [15]
                  ,SIIAPP_Bod.Cap_Sum AS [16]
                  ,SIIAPP_Bod.Cap_Color AS [17]
                  ,SIIAPP_Bod.Cap_Material AS [18]
                  ,SIIAPP_Bod.Foil_Cod AS [19]
                  ,SIIAPP_Bod.Foil AS [20]
                  ,SIIAPP_Bod.Foil_Type AS [21]
                  ,SIIAPP_Bod.FB_Cod AS [22]
                  ,SIIAPP_Bod.Funda_Banda AS [23]
                  ,SIIAPP_Bod.Ubicacion_Termoduc AS [24]
                  ,SIIAPP_Bod.Etiq_Cod AS [25]
                  ,SIIAPP_Bod.Etiqueta AS [26]
                  ,SIIAPP_Bod.Box_Cod AS [27]
                  ,SIIAPP_Bod.Box_Folding AS [28]
                  ,SIIAPP_Bod.Box_Sum AS [29]
                  ,SIIAPP_Cal.Mesofilos AS [30]
                  ,SIIAPP_Cal.E_Coli AS [31]
                  ,SIIAPP_Cal.S_Aureus AS [32]
                  ,SIIAPP_Cal.P_Aureoginosa AS [33]
                  ,SIIAPP_Cal.Moho_Levadura AS [34]
                  ,SIIAPP_Cal.Micro_Obs AS [35]
                  ,SIIAPP_Cal.Cumple_Microbiologia AS [36]
                  ,SIIAPP_Cal.Densidad AS [37]
                  ,SIIAPP_Cal.Ph AS [38]
                  ,SIIAPP_Cal.Contenido AS [39]
                  ,SIIAPP_Cal.Color AS [40]
                  ,SIIAPP_Cal.Textura AS [41]
                  ,SIIAPP_Cal.Olor AS [42]
                  ,SIIAPP_Cal.Viscosidad AS [43]
                  ,SIIAPP_Cal.Rpm AS [44]
                  ,SIIAPP_Cal.Aguja AS [45]
                  ,SIIAPP_Cal.Torque AS [46]
                  ,SIIAPP_Prod.Cod_Caja AS [47]
                  ,SIIAPP_Prod.Caja_Embalaje AS [48]
                  ,SIIAPP_Prod.FB_Dimensiones_Adec AS [49]
                  ,SIIAPP_Prod.FB_Tunel_Temp AS [50]
                  ,SIIAPP_Prod.FB_Tunel_Speed AS [51]
                  ,SIIAPP_Prod.FB_Test_Tool AS [52]
                  ,SIIAPP_Prod.FB_obs AS [53]
                  ,SIIAPP_Prod.SCE_Temp AS [54]
                  ,SIIAPP_Prod.SCE_Temp_Time AS [55]
                  ,SIIAPP_Prod.SCE_Obs AS [56]
                  ,SIIAPP_Prod.SCE_Hermt AS [57]
                  ,in_items.itetipoitem AS [58]
                  ,SIIAPP_Des.NIT AS [59]
                  ,SubQuery.prvversion AS [60]
                  ,SIIAPP_Des.Comercial AS [61]
                  ,SIIAPP_Cal.Titular_Nsoc AS [62]
                  ,SIIAPP_Cal.Uso_Exclu AS [63]
                  ,SIIAPP_Bod.Bottle_Type AS [64]
                  ,SIIAPP_Bod.Bottle_Cuello AS [65]
                  ,SIIAPP_Bod.Bottle_Sellado AS [66]
                  ,SIIAPP_Bod.Cap_Type AS [67]
                  ,SIIAPP_Bod.Cap_Valvule AS [68]
                  ,SIIAPP_Bod.Cap_Overcap AS [69]
                  ,SIIAPP_Bod.Cap_Style AS [70]
                  ,SIIAPP_Bod.Cap_Casquete_End AS [71]
                  ,SIIAPP_Bod.Cap_Casquete_Color AS [72]
                  ,SIIAPP_Prod.Box_Units AS [73]
                  ,SIIAPP_Prod.Comb_Dimensions AS [74]
                  ,SIIAPP_Prod.Box_Print_Texture AS [75]
                  ,SIIAPP_Cal.Forma_Cosm AS [76]
                  ,SIIAPP_Bod.Bottle_Deco AS [77]
                  ,SIIAPP_Bod.Bottle_Deco_Cod AS [78]
                  ,SIIAPP_Bod.Bottle_Deco_Desc AS [79]
                  ,SIIAPP_Bod.Cap_Deco AS [80]
                  ,SIIAPP_Bod.Cap_Deco_Cod AS [81]
                  ,SIIAPP_Bod.Cap_Deco_Desc AS [82]
                  ,SIIAPP_Bod.Cap_Deco_Etiq AS [83]
                  ,SIIAPP_Bod.Cap_Deco_Etiq_Cod AS [84]
                  ,SIIAPP_Bod.Foil_Bool AS [85]
                  ,SIIAPP_Bod.FB_Bool AS [86]
                  ,SIIAPP_Bod.Escurr_Bool AS [87]
                  ,SIIAPP_Bod.Escurr_Cod AS [88]
                  ,SIIAPP_Bod.Escurr_Desc AS [89]
                  ,SIIAPP_Bod.Capil_Bool AS [90]
                  ,SIIAPP_Bod.Capil_Obs AS [91]
                  ,SIIAPP_Bod.Capil_Dimen AS [92]
                  ,SIIAPP_Bod.Acces_Bool AS [93]
                  ,SIIAPP_Bod.Acces_Which AS [94]
                  ,SIIAPP_Prod.Box_Corr AS [95]
                  ,SIIAPP_Prod.Comb_Bool AS  [96]
                  ,SIIAPP_Prod.Rotulo AS [97]
                  ,SIIAPP_Prod.Rotulo_text AS [98]
                  ,SIIAPP_Prod.Comb_key AS [99]
                  ,SIIAPP_Des.Marca AS [100]
                  FROM SIIAPP.dbo.SIIAPP_Bod
                  INNER JOIN SIIAPP.dbo.SIIAPP_Des
                    ON SIIAPP_Bod.N_Control = SIIAPP_Des.N_control
                  INNER JOIN SIIAPP.dbo.SIIAPP_Cal
                    ON SIIAPP_Cal.N_Control = SIIAPP_Des.N_control
                  INNER JOIN SIIAPP.dbo.SIIAPP_Prod
                    ON SIIAPP_Prod.N_Control = SIIAPP_Des.N_control
                  LEFT OUTER JOIN ssf_genericos.dbo.in_items
                    ON in_items.itecodigo = SIIAPP_Des.PT
                  LEFT OUTER JOIN (SELECT
                      pd_procesoversion.prvcodiitem
                    ,pd_procesoversion.prvcompania
                    ,pd_procesoversion.prvversion
                    FROM ssf_genericos.dbo.pd_procesoversion
                    WHERE pd_procesoversion.prvdefecto = 'S') SubQuery
                    ON SubQuery.prvcodiitem = in_items.itecodigo
                      AND SubQuery.prvcompania = in_items.itecompania
                  WHERE SIIAPP_Bod.N_Control = ?


                                  """                    