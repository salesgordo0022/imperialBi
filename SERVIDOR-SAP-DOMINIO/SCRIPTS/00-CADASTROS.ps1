<#
.SYNOPSIS
    Extrai os cadastros que TODOS os relatórios usam.

.DESCRIPTION
    Empresas, CNAE 2.0, plano de contas, contas de caixa e usuários.
    Roda em poucos segundos. Os scripts 01 a 06 chamam este sozinhos
    quando os arquivos ainda não existem — você raramente precisa
    rodá-lo à mão.

.EXAMPLE
    .\00-CADASTROS.ps1
    .\00-CADASTROS.ps1 -Forcar     # reextrai mesmo se já existir
#>
[CmdletBinding()]
param([switch]$Forcar)

. "$PSScriptRoot\_lib\dominio.ps1"

$arquivos = @('01_empresas', '01b_cnae', '24_cnae20', '20_plano_contas', '21_contas_caixa', '23_usuarios')

if (-not $Forcar) {
    $faltando = $arquivos | Where-Object { -not (Test-Path (Join-Path $DOM.PastaDados "$_.tsv")) }
    if (-not $faltando) { return }
}

Start-DomRelatorio "CADASTROS (base comum a todos os relatórios)"

Export-DomTabela -Nome '01_empresas' `
    -Header "codi_emp`tnome_emp`trazao_emp`tcgce_emp`tiest_emp`timun_emp`tcnae_emp`tramo_emp`tcida_emp`testa_emp`tstat_emp`ttins_emp`tsimples_emp`tuefi_emp`tucta_emp`tufol_emp`tucxa_emp`tupat_emp`tulal_emp`tuhon_emp`tcnpj_prod_rural`tdcad_emp`tdina_emp" `
    -Sql @"
SELECT codi_emp, nome_emp, razao_emp, cgce_emp, iest_emp, imun_emp, cnae_emp, ramo_emp,
       cida_emp, esta_emp, stat_emp, tins_emp, simples_emp,
       uefi_emp, ucta_emp, ufol_emp, ucxa_emp, upat_emp, ulal_emp, uhon_emp,
       CNPJ_PRODUTOR_RURAL_EMP, dcad_emp, dina_emp
FROM bethadba.geempre
"@

# cnae_emp (numérico) vive quase sempre vazio; o CNAE bom é i_cnae20
Export-DomTabela -Nome '01b_cnae' -Header "codi_emp`ti_cnae20`tramo_emp" -Sql @"
SELECT codi_emp, i_cnae20, ramo_emp FROM bethadba.geempre
"@

Export-DomTabela -Nome '24_cnae20' -Header "codigo`tdescricao" -Sql @"
SELECT i_cnae20, descricao FROM bethadba.gecnae20
"@

Export-DomTabela -Nome '20_plano_contas' `
    -Header "codi_emp`tcodi_cta`tnome_cta`tclas_cta`ttipo_cta`tsituacao_cta`tcarne_leao_cta`tlcdpr_cta`tgrdre_cta" `
    -Sql @"
SELECT codi_emp, codi_cta, nome_cta, clas_cta, tipo_cta, SITUACAO_CTA,
       CARNE_LEAO_CTA, LCDPR_CTA, grdre_cta
FROM bethadba.ctcontas
"@

Export-DomTabela -Nome '21_contas_caixa' -Header "codi_emp`tcodi_cta`tsele_cta`ttipo_conta_caixa" -Sql @"
SELECT codi_emp, codi_cta, sele_cta, TIPO_CONTA_CAIXA FROM bethadba.ctcontacaixa
"@

# geusuarios está vazia nesta base; os usuários de verdade estão aqui
Export-DomTabela -Nome '23_usuarios' -Header "id`tnome`tgrupo" -Sql @"
SELECT I_SECUSUARIOS, NOME, GRUPO FROM bethadba.GESECUSUARIOS
"@

Write-Host "   cadastros prontos." -ForegroundColor Green
