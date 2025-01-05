--------------------------------------------------------------------------------
--	uspAttendantFutureFlights
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspAttendantFutureFlights(
     @intAttendant_ID AS INTEGER
)
AS
BEGIN
Select TF.intFlightID ,TF.dtmFlightDate, TF.strFlightNumber, TF.dtmTimeofDeparture, TF.dtmTimeOfLanding, TFAP.strAirportCity AS FromAirport, TTAP.strAirportCity AS ToAirport, TF.intMilesFlown, TPL.strPlaneNumber 
 From 
 TFlights	    as TF JOIN TAttendantFlights  as TAF
 ON TF.intFlightID = TAF.intFlightID
 JOIN TAttendants as TA
 ON TA.intAttendantID = TAF.intAttendantID
 JOIN TAirports as TFAP
 ON TFAP.intAirportID = TF.intFromAirportID
 JOIN TAirports as TTAP
 ON TTAP.intAirportID = TF.intToAirportID
 JOIN TPlanes as TPL
 ON TPL.intPlaneID = TF.intPlaneID
 WHERE TA.intAttendantID = @intAttendant_ID and TF.dtmFlightDate > GetDate()
END;
-- --------------------------------------------------------------------------------
--	uspAttendantPastFlights
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspAttendantPastFlights(
     @intAttendant_ID AS INTEGER
    
)
AS
BEGIN
Select TF.intFlightID ,TF.dtmFlightDate, TF.strFlightNumber, TF.dtmTimeofDeparture, TF.dtmTimeOfLanding, TFAP.strAirportCity AS FromAirport, TTAP.strAirportCity AS ToAirport, TF.intMilesFlown, TPL.strPlaneNumber 
 From 
 TFlights	    as TF JOIN TAttendantFlights  as TAF
 ON TF.intFlightID = TAF.intFlightID
 JOIN TAttendants as TA
 ON TA.intAttendantID = TAF.intAttendantID
 JOIN TAirports as TFAP
 ON TFAP.intAirportID = TF.intFromAirportID
 JOIN TAirports as TTAP
 ON TTAP.intAirportID = TF.intToAirportID
 JOIN TPlanes as TPL
 ON TPL.intPlaneID = TF.intPlaneID
 WHERE TA.intAttendantID = 1 and TF.dtmFlightDate <= GetDate()
END;
-- --------------------------------------------------------------------------------
--	uspPilotFutureFlights
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspPilotFutureFlights(
     @intPilot_ID AS INTEGER
    
)
AS
BEGIN
Select TF.intFlightID ,TF.dtmFlightDate, TF.strFlightNumber, TF.dtmTimeofDeparture, TF.dtmTimeOfLanding, TFAP.strAirportCity AS FromAirport, TTAP.strAirportCity AS ToAirport, TF.intMilesFlown, TPL.strPlaneNumber 
                         From 
                         TFlights	    as TF JOIN TPilotFlights  as TPOF
	                     ON TF.intFlightID = TPOF.intFlightID
	                     JOIN TPilots as TPO
	                     ON TPO.intPilotID = TPOF.intPilotID
                         JOIN TAirports as TFAP
                         ON TFAP.intAirportID = TF.intFromAirportID
                         JOIN TAirports as TTAP
                         ON TTAP.intAirportID = TF.intToAirportID
                         JOIN TPlanes as TPL
                         ON TPL.intPlaneID = TF.intPlaneID
						 WHERE TPO.intPilotID = @intPilot_ID and TF.dtmFlightDate > GetDate()
END;
-- --------------------------------------------------------------------------------
--	uspPilotPastFlights
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspPilotPastFlights(
     @intPilot_ID AS INTEGER
    
)
AS
BEGIN
Select TF.intFlightID ,TF.dtmFlightDate, TF.strFlightNumber, TF.dtmTimeofDeparture, TF.dtmTimeOfLanding, TFAP.strAirportCity AS FromAirport, TTAP.strAirportCity AS ToAirport, TF.intMilesFlown, TPL.strPlaneNumber 
                         From 
                         TFlights	    as TF JOIN TPilotFlights  as TPOF
	                     ON TF.intFlightID = TPOF.intFlightID
	                     JOIN TPilots as TPO
	                     ON TPO.intPilotID = TPOF.intPilotID
                         JOIN TAirports as TFAP
                         ON TFAP.intAirportID = TF.intFromAirportID
                         JOIN TAirports as TTAP
                         ON TTAP.intAirportID = TF.intToAirportID
                         JOIN TPlanes as TPL
                         ON TPL.intPlaneID = TF.intPlaneID
						 WHERE TPO.intPilotID = @intPilot_ID and TF.dtmFlightDate <= GetDate()
END;
-- --------------------------------------------------------------------------------
--	uspTotalFlights
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspTotalFlights
AS
BEGIN
SELECT COUNT(TFP.intFlightPassengerID) as TotalFlights
                        FROM TFlightPassengers as TFP
END;
-- --------------------------------------------------------------------------------
--	uspTotalPassengersInDB
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspTotalPassengersInDB
AS
BEGIN
SELECT COUNT(TP.intPassengerID) as TotalPassengers
                        From TPassengers as TP
END;
-- --------------------------------------------------------------------------------
--	uspAddPassenger
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspAddPassenger
     @intPassengerID				AS INTEGER OUTPUT
    ,@strFirstName				AS VARCHAR(255)
    ,@strLastName				AS VARCHAR(255)
    ,@strAddress				AS VARCHAR(255)
    ,@strCity					AS VARCHAR(255) 
    ,@intState					AS INTEGER 
    ,@strZip					AS VARCHAR(255)
    ,@strPhoneNumber			AS VARCHAR(255)
    ,@strEmail					AS VARCHAR(255)
	,@strPassengerLoginID		AS VARCHAR(255)
	,@strPassengerPassword		AS VARCHAR(255)
	,@dtmPassengerDateOfBirth	AS DATETIME
       
AS
SET XACT_ABORT ON --terminate and rollback if any errors
BEGIN TRANSACTION
    SELECT @intPassengerID = MAX(intPassengerID) + 1 
    FROM TPassengers (TABLOCKX) -- lock table until end of transaction
    -- default to 1 if table is empty
    SELECT @intPassengerID = COALESCE(@intPassengerID, 1)
    INSERT INTO TPassengers (intPassengerID, strFirstName, strLastName, strAddress, strCity, intStateID, strZip, strPhoneNumber, strEmail, strPassengerLoginID, strPassengerPassword, dtmPassengerDateOfBirth)
    VALUES (@intPassengerID, @strFirstName, @strLastName, @strAddress, @strCity, @intState, @strZip, @strPhoneNumber, @strEmail, @strPassengerLoginID, @strPassengerPassword, @dtmPassengerDateOfBirth)

COMMIT TRANSACTION
GO
-- --------------------------------------------------------------------------------
--	uspUpdatePassenger
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspUpdatePassenger
     @intPassengerID				AS INTEGER OUTPUT
    ,@strFirstName				AS VARCHAR(255)
    ,@strLastName				AS VARCHAR(255)
    ,@strAddress				AS VARCHAR(255)
    ,@strCity					AS VARCHAR(255) 
    ,@intState					AS INTEGER 
    ,@strZip					AS VARCHAR(255)
    ,@strPhoneNumber			AS VARCHAR(255)
    ,@strEmail					AS VARCHAR(255)
	,@strPassengerLoginID		AS VARCHAR(255)
	,@strPassengerPassword		AS VARCHAR(255)
	,@dtmPassengerDateOfBirth	AS DATETIME
AS
SET XACT_ABORT ON --terminate and rollback if any errors
BEGIN TRANSACTION
Update TPassengers

	SET strFirstName = @strFirstName,		
	strLastName	= @strLastName,		
	strAddress = @strAddress,		
	strCity = @strCity,			
	intStateID = @intState,			
	strZip = @strZip,					
	strPhoneNumber = @strPhoneNumber,			
	strEmail = @strEmail,
	strPassengerLoginID = @strPassengerLoginID,
	strPassengerPassword = @strPassengerPassword,
	dtmPassengerDateOfBirth = @dtmPassengerDateOfBirth
		WHERE  intPassengerID = @intPassengerID
COMMIT TRANSACTION
GO

-- --------------------------------------------------------------------------------
--	uspDeletePilot
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspDeletePilot
     @intPilotID				AS INTEGER  
    
       
AS
SET XACT_ABORT ON --terminate and rollback if any errors
BEGIN TRANSACTION
  
    Delete  FROM TPilots 
	WHERE  intPilotID = @intPilotID
	
COMMIT TRANSACTION
GO
-- --------------------------------------------------------------------------------
--	uspAddFlight
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspAddFlight
     @intFlightID				AS INTEGER OUTPUT
    ,@dtmFlightDate				DATETIME
	,@strFlightNumber			AS VARCHAR(255)
	,@dtmTimeofDeparture			DATETIME
    ,@dtmTimeOfLanding			DATETIME
    ,@intFromAirportID			AS INTEGER
    ,@intToAirportID			AS INTEGER 
    ,@intMilesFlown				AS INTEGER
    ,@intPlaneID			AS INTEGER

       
AS
SET XACT_ABORT ON --terminate and rollback if any errors
BEGIN TRANSACTION
    SELECT @intFlightID = MAX(intFlightID) + 1 
    FROM TFlights (TABLOCKX) -- lock table until end of transaction
    -- default to 1 if table is empty
    SELECT @intFlightID = COALESCE(@intFlightID, 1)
    INSERT INTO TFlights (intFlightID, dtmFlightDate, strFlightNumber, dtmTimeofDeparture, dtmTimeOfLanding, intFromAirportID, intToAirportID, intMilesFlown, intPlaneID)
    VALUES (@intFlightID, @dtmFlightDate, @strFlightNumber, @dtmTimeofDeparture, @dtmTimeOfLanding, @intFromAirportID, @intToAirportID, @intMilesFlown, @intPlaneID)

COMMIT TRANSACTION
GO
-- --------------------------------------------------------------------------------
--	uspPLogin
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspPLogin(
      @strPassengerLoginID AS VARCHAR(255)
     ,@strPassengerPassword AS VARCHAR(255)
)
AS
BEGIN
Select TP.intPassengerID
From TPassengers as TP
WHERE TP.strPassengerLoginID = @strPassengerLoginID
AND  TP.strPassengerPassword = @strPassengerPassword
END;
-- --------------------------------------------------------------------------------
--	uspELogin
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspELogin(
       @strEmployeeLoginID AS VARCHAR(255)
		,@strEmployeePassword AS VARCHAR(255)
)
AS
BEGIN
SELECT TE.strEmployeeRole, TE.intEmployeePK
FROM TEmployees as TE
WHERE TE.strEmployeeLoginID = @strEmployeeLoginID
AND TE.strEmployeePassword = @strEmployeePassword
END;
-- --------------------------------------------------------------------------------
--	uspUpdateEmployeeLogin
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspUpdateEmployeeLogin
	 @strEmployeeRole				AS VARCHAR(255) OUTPUT
	,@intEmployeePK					AS INTEGER OUTPUT
	,@strEmployeeLoginID			AS VARCHAR(255)
    ,@strEmployeePassword			AS VARCHAR(255)
AS
SET XACT_ABORT ON --terminate and rollback if any errors
BEGIN TRANSACTION
Update TEmployees 
	SET strEmployeeRole = @strEmployeeRole,
	intEmployeePK = @intEmployeePK,
	strEmployeeLoginID = @strEmployeeLoginID,
	strEmployeePassword = @strEmployeePassword
	Where strEmployeeRole = @strEmployeeRole
	AND intEmployeePK = @intEmployeePK
COMMIT TRANSACTION
-- --------------------------------------------------------------------------------
--	uspLoginInfo
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspLoginInfo(
     @strEmployeeRole				AS VARCHAR(255)
	,@intEmployeePK					AS INTEGER 
)
AS
BEGIN
	Select TE.strEmployeeLoginID, TE.strEmployeePassword
	From TEmployees as TE
	Where strEmployeeRole = @strEmployeeRole
	AND intEmployeePK = @intEmployeePK
END;
------------------------------------------------------------------------------
--	uspMilesFlown
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspMilesFlown(
     @intFlight_ID AS INTEGER
    
)
AS
BEGIN
Select TF.intMilesFlown
FROM TFlights as TF
Where TF.intFlightID = @intFlight_ID
END;
-- --------------------------------------------------------------------------------
--	uspTotalPassengersPerFlight
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspTotalPassengersPerFlight(
     @intFlight_ID AS INTEGER
    
)
AS
BEGIN
Select COUNT(TFP.intFlightPassengerID) as TotalPassengersPerFlight
FROM TFlightPassengers as TFP JOIN TFlights as TF
ON TF.intFlightID = TFP.intFlightID
Where TF.intFlightID = @intFlight_ID
END;
-- --------------------------------------------------------------------------------
--	uspTypeOfPlane
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspTypeOfPlane(
     @intFlight_ID AS INTEGER
    
)
AS
BEGIN
Select TPTY.intPlaneTypeID
FROM TPlaneTypes as TPTY JOIN TPlanes as TPL
ON TPTY.intPlaneTypeID = TPL.intPlaneTypeID
JOIN TFlights as TF
ON TPL.intPlaneID = TF.intPlaneID
Where TF.intFlightID = @intFlight_ID
END;
-- --------------------------------------------------------------------------------
--	uspAirPlaneDestination
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspAirPlaneDestination(
     @intFlight_ID AS INTEGER
    
)
AS
BEGIN
Select TF.intToAirportID as Destination
FROM TFlights as TF
Where TF.intFlightID = @intFlight_ID
END;
-- --------------------------------------------------------------------------------
--	uspPassengerDay
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspPassengerDay(
     @intPassenger_ID AS INTEGER
    
)
AS
BEGIN
Select DATEDIFF(DAY, TP.dtmPassengerDateOfBirth, GETDATE()) as PassengerDay
FROM TPassengers as TP
Where TP.intPassengerID = @intPassenger_ID
END;
-- --------------------------------------------------------------------------------
--	uspRepeatCustomer
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspRepeatCustomer(
     @intPassenger_ID AS INTEGER
    
)
AS
BEGIN
Select Count(TFP.intPassengerID) as RepeatCustomer
From TFlightPassengers as TFP
Where TFP.intPassengerID = @intPassenger_ID
END;

-- --------------------------------------------------------------------------------
--	uspAddEmployeeLogin
-- --------------------------------------------------------------------------------
CREATE PROCEDURE uspAddEmployeeLogin
     @intEmployeeID					AS INTEGER OUTPUT
	,@strEmployeeLoginID			AS VARCHAR(255)
	,@strEmployeePassword			AS VARCHAR(255)
	,@strEmployeeRole				AS VARCHAR(255) 
	,@intEmployeePK					AS INTEGER 
    
       
AS
SET XACT_ABORT ON --terminate and rollback if any errors
BEGIN TRANSACTION
    SELECT @intEmployeeID = MAX(intEmployeeID) + 1
    FROM TEmployees (TABLOCKX) -- lock table until end of transaction
    -- default to 1 if table is empty
    SELECT @intEmployeeID = COALESCE(@intEmployeeID, 1)
    INSERT INTO TEmployees (intEmployeeID, strEmployeeLoginID, strEmployeePassword, strEmployeeRole, intEmployeePK)
    VALUES (@intEmployeeID, @strEmployeeLoginID, @strEmployeePassword, @strEmployeeRole, @intEmployeePK)

COMMIT TRANSACTION
GO
