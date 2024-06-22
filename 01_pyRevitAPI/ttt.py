def createNewTee(pipe1, pipe2, pipe3):
	def closetConn(Pipe1, Pipe2, Pipe3):
		connectors1 = list(Pipe1.ConnectorManager.Connectors.GetEnumerator())
		connectors2 = list(Pipe2.ConnectorManager.Connectors.GetEnumerator())
		connectors3 = list(Pipe3.ConnectorManager.Connectors.GetEnumerator())
		Connector1 = NearestConnector(connectors1, Pipe2)
		Connector2 = NearestConnector(connectors2, Pipe1)
		Connector3 = NearestConnector(connectors3, Pipe1)
		return Connector1, Connector2, Connector3
	closetConn = closetConn(pipe1 , pipe2, pipe3)
	fittings = []
	TransactionManager.Instance.EnsureInTransaction(doc)
	fitting = doc.Create.NewTeeFitting(closetConn[0], closetConn[1], closetConn[2])					
	fitting_dynamo = fitting.ToDSType(False)
	fittings.append(fitting_dynamo)
	TransactionManager.Instance.TransactionTaskDone()
	return fittings