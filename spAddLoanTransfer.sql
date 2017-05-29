CREATE PROCEDURE spAddLoanTransfer

@traAccountCode	VARCHAR(50),
@traModelNo		VARCHAR(50),
@traSerialNo	VARCHAR(50),
@traConnote		VARCHAR(50),
@traRecipient	VARCHAR(50),
@traCreatedBy  	VARCHAR(50)

AS
BEGIN
	-- Check if the item already exists
	IF EXISTS(SELECT * FROM tbl_loan_transfer 
					WHERE traAccountCode = @traAccountCode 
						AND traModelNo = @traModelNo 
						AND traSerialNo = @traSerialNo 
						AND traRecipient = @traRecipient)
		BEGIN
			RAISERROR('This item already exists in the Loan Transfer table.', 12, 12)
			RETURN(@@error)
		END
	ELSE

	BEGIN
		INSERT INTO tbl_loan_transfer (
			traAccountCode, 
			traModelNo, 
			traSerialNo, 
			traConnote, 			
			traRecipient,
			traCreatedBy)
		VALUES (
			@traAccountCode,
			@traModelNo,
			@traSerialNo,
			@traConnote,			
			@traRecipient,
			@strCreatedBy)

		IF @@ERROR <> 0 
			BEGIN
				RAISERROR('An error occured while trying to add this item.', 12, 12)
				RETURN(@@error)
			END
	END
END
GO
