<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Controle de Processos</title>
</head>
<body>
    <%
        'Recuperando dados do formulário
        nome    =Request.Form("nome")
        cpf     =Request.Form("cpf")
        convenio    =Request.Form("convenio")
        usuario     =Request.Form("convenio")
        senha   =Request.Form("senha")
        senha2  =Request.Form("senha2")

        'Limpando lixos e previnindo ataques removendo espaços em brancos e aspas simples
        nome    =LUCASE(REPLACE(TRIM(name)"'",""))
        cpf     =REPLACE(TRIM(cpf)"'","")
        convenio    =REPLACE(TRIM(convenio)"'","")
        usuario     =LUCASE(REPLACE(TRIM(usuario)"'",""))
        senha   =REPLACE(TRIM(senha)"'","")
        senha2  =REPLACE(TRIM(senha2)"'","")

        'Validando Formulario
        'Verificando se o campo convenio tem apenas números
        if isNumeric(convenio) = False Then
            ValidaConvenio = "erro"
            OcorreuErro = "sim"
        end if

        'Validando o campo senha 
        'Verificando se a senha digitada possui no minino 6 e no maximo 12
        if len(senha)<6 or len(senha2) <6 or len (senha)>10 or len(senha2)<10 Then
            ValidaSenha = "erro"
            OcorreuErro = "Sim"
        end if

        'Verificando se as senhas são iguais
        if senha <> senha2 Then
            SenhaIguais = "erro"
            OcorreuErro = "sim"
        end if

        'Localizando usuario no banco para inibir duplicatas
        StrConveniado = "SELECT Nome_Conveniado, Usuario, CPF_Convenio FROM Conveniados Where Nome_Conveniados = '" & nome & "'  OR
        Usuario = '" & usuario & "' OR CPF_Conveniado = '" & cpf & "'"
        Set rsSQL = conexao.Execute(StrConveniado)

        if not rsSQL.EOF then
            ValidaUsuariosDiferentes = "erro"
            OcorreuErro = "sim"
        
        else
            if(ValidaConvenio <> "erro") and (ValidaTamanhoSenha <> "erro") and (SenhasIguais <> "erro") then
                Response.Write "Não ocorreram erros no formulário" & "<br>"

                Set add_action = Server.CreateObject("ADODB.Recordset")
                
                add_action.Open "Conveniados", conexao, 3,3

                add_action.AddNew

                    add_action("Nome_Conveniado")   = nome
                    add_action("CPF_Conveniados")   = cpf
                    add_action("Convenios")         = convenio
                    add_action("Usuario")           = usuario
                    add_action("Senha")             = Senha
            

                add_action.Update

                add_action.Close
                conexao.Close

                Set add_action = Nothing
                Set conexao = Nothing
            
            end if
        end if
    %>
    <%  
        
        if ValidaConvenio = "erro" then
            Response.Write "<font-style= 'color:red;'> O campo Convenio deve conter apenas números!" & "</font><br>"
        end if

        if ValidaTamanhoSenha = "erro" then
            Response.Write "<font style='color: red;'>Sua senha de ter no mínimo 6 e no maximo 10 caracteres!" & "</font><br>"
        end if

        if ValidaSenhaIguais = "erro" then
            Response.Write "<font style= 'color: red;' >Senhas não conferem!" & "</font><br>"
        end if

        if ValidaUsuariosDiferentes = "erro"then
            Response.Write "<font style= color: red;'>Existe um Conveniado com o nome: " & nome & "</font><br>"
            Response.Write "<font style= color: red;'>Ou existe um usuario: " & usuario & "</font><br>"
            Response.Write "<font style= color: red;'>Ou esse cpf ja existe: " & cpf & "</font><br>"
        end if
        if OcorreuErro = "sim" then
    %>
            <p align="center">Conveniados[<small>Inclusão</small></p>

            <div style="text-align: left;position: relative; left: 50%; top: 10%; margin-left: -110px; margin-top: 30px;">
                <form class="form" action="Conveniados_inclusao.html" method="post">
                    <p align="left">
                        <label for="nome">Nome Completo:</label><br>
                        <input class="nome" type="text" required="required"  name="nome" autocomplete="on" id = nome placeholder="Nome Completo" size="35px" minlength="6" maxlength="20" autofocus value = "<%=nome%>" >
                    </p>
                    <p>
                        <label for="cpf">CPF:</label><br>
                        <input class="cpf" type="text" name="cpf" required title="Preencha o campo apenas com números" id = cpf placeholder="Apenas Números" size="35px" pattern="\d{11}" minlength="11" value = "<%=cpf%>" >
                    </p>
                    <p>
                        <label for="convenio">Número do Convênio:</label><br>
                        <input class="convenio" type="text" name="convenio" required id="convenio" placeholder="Número do Convênio" size="35px" pattern="\d{10}" maxlength="10" value = "<%=convenio%>">
                    </p>
                    <p>
                        <label for="usuario">Usuário:</label><br>
                        <input class="usuario" type="text" name="usuario" id="usuario" placeholder="Nome de Usuário" size="35px" required value = "<%=usuario%>">
                    </p>
                    <p>
                        <label for="senha">Senha:</label><br>
                        <input class="senha" type="password" name="senha" id="senha" placeholder="Digite sua senha" size="35px" required minlength="6" maxlength="12">
                    </p>
                    <p>
                        <label for="senha2">Confirme sua senha:</label><br>
                        <input class="senha2" type="password" name="senha2" id="senha2" placeholder="" size="35px" required>
                    </p>
    
                    <input class="submit" type="submit" name="submit" id="submit">
                    <input class="reset" type="reset" name="reset" id="reset">
                </form>
            </div>
    <%
        end if
    %>
</body>
</html>