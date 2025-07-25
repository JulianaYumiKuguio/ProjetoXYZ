SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[ObterTransacoesCategorizadasPorPeriodo]
(
    @DataInicio DATE, -- Data de início do período
    @DataFim DATE     -- Data de fim do período
)
RETURNS TABLE
AS
RETURN
(
    SELECT
        T.ID_Transacao,
		T.Numero_Cartao,
        T.Data_Transacao,
        T.Valor_Transacao,
        T.Descricao,
		CASE
		WHEN Status = 1 THEN 'Aprovada' 
		WHEN Status = 2 THEN 'Pendente' 
		ELSE 'Cancelada' END
	as Status,
        dbo.ObterCategoriaValor(T.Valor_Transacao) AS CategoriaValor -- Chama a Scalar Function aqui!
    FROM
        tb_Transacoes AS T
    WHERE
        T.Data_Transacao >= @DataInicio
        AND T.Data_Transacao <= @DataFim -- Use <= para incluir o último dia
);