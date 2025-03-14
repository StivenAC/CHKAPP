CREATE TRIGGER dbo.trigger1
ON dbo.SIIAPP_Des
AFTER INSERT
AS
BEGIN
    -- Ensure a transactional scope for all operations
    SET NOCOUNT ON;

    -- Insert the new N_Control value into SIIAPP_Bod
    INSERT INTO dbo.SIIAPP_Bod (N_Control)
    SELECT N_Control
    FROM INSERTED;

    -- Insert the new N_Control value into SIIAPP_Prod
    INSERT INTO dbo.SIIAPP_Prod (N_Control)
    SELECT N_Control
    FROM INSERTED;

    -- Insert the new N_Control value into SIIAPP_Cal
    INSERT INTO dbo.SIIAPP_Cal (N_Control)
    SELECT N_Control
    FROM INSERTED;
END;
GO