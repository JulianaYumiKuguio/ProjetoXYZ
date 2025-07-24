
/****** Object:  View [dbo].[VW_ConsultaTransacoesPorCategoria]    Script Date: 24/07/2025 18:30:12 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





CREATE VIEW [dbo].[VW_ConsultaTransacoesPorCategoria] AS


	select 
	Id_Transacao as ID, 
	CAST(Numero_Cartao AS VARCHAR(16)) AS Numero_Cartao,
	Valor_Transacao, 
	Data_Transacao,
	Descricao, 
	dbo.ObterCategoriaValor(valor_transacao) as Categoria,
	CASE
		WHEN Status = 1 THEN 'Aprovada' 
		WHEN Status = 2 THEN 'Pendente' 
		ELSE 'Cancelada' END
	as Status
	from [dbo].[tb_Transacoes]



GO


