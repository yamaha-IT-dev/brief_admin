CREATE PROCEDURE spAddBasket

@strItemCode		VARCHAR(50),
@strSerialNo		VARCHAR(50),
@strLIC			VARCHAR(50),
@strAccountCode		VARCHAR(9),
@intOrderNo		INT,
@intOrderLine		INT

AS

BEGIN

	-- We first have to check if the item already exists for the particular account
	IF EXISTS(SELECT * FROM workflow_loan_return_item_list WHERE item_code = @strItemCode AND account_code = @strAccountCode AND order_no = @intOrderNo AND order_line = @intOrderLine)
		BEGIN
			RAISERROR('This item already exists for this account.', 12, 12)
			RETURN(@@error)
		END
	ELSE

	BEGIN
		INSERT INTO workflow_loan_return_item_list (
			item_code, 
			serial_number, 
			product_lic, 
			account_code, 
			order_no, 
			order_line)
		VALUES (
			@strItemCode,
			@strSerialNo,
			@strLIC,
			@strAccountCode,
			@intOrderNo,
			@intOrderLine)

		IF @@ERROR <> 0 
			BEGIN
				RAISERROR('An error occured while trying to add this item.', 12, 12)
				RETURN(@@error)
			END
	END
END
GO
