package main

import (
	"fmt"
	"log"
	"strings"

	"github.com/xuri/excelize/v2"
)

func criarPlanilha() {

	// cria um novo arquivo
	f := excelize.NewFile()

	sheet := "Vendas"
	f.SetSheetName("Sheet1", sheet)

	// cabeçalho da planilha
	headers := []string{
		"Nome do Provedor",
		"Cidade",
		"Estado",
		"Telefone",
		"Nome do Contato",
		"Data do Primeiro Contato",
		"Status",
	}

	// preencher os cabeçalhos
	for i, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheet, cell, header)
	}

	// Estilo cabeçalho
	style, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#D9D9D9"},
			Pattern: 1,
		},
	})
	if err != nil {
		fmt.Println("Erro ao criar estilo:", err)
		return
	}

	f.SetCellStyle(sheet, "A1", "G1", style)

	// ajusta largura das colunas
	f.SetColWidth(sheet, "A", "A", 25)
	f.SetColWidth(sheet, "B", "C", 15)
	f.SetColWidth(sheet, "D", "D", 18)
	f.SetColWidth(sheet, "E", "E", 22)
	f.SetColWidth(sheet, "F", "G", 20)

	// Salva o arquivo
	err = f.SaveAs("controle_vendas_provedores.xlsx")
	if err != nil {
		fmt.Println("Erro ao salvar arquivo:", err)
		return
	}

	fmt.Println("Planilha criada com sucesso!!!")

}

func lerPlanilha() {

	// abre planilha existente
	f, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		log.Fatal("Erro ao abrir a planilha", err)
	}

	// lê todas as linhas da aba Vendas
	rows, err := f.GetRows("Vendas")
	if err != nil {
		log.Fatal("Erro ao ler linhas da aba Vendas", err)
	}

	fmt.Println("Contatos com status NOVO:")
	fmt.Println("-------------------------")

	contadorNovo := 0

	// percorre as linhas ignorando o cabeçalho
	for i, row := range rows {
		if i == 0 {
			continue // pula cabeçalho
		}

		// verifica se linha tem pelo menos 7 colunas
		if len(row) >= 7 {
			if len(row) >= 7 {
				status := strings.TrimSpace(row[6])
				statusNormalizado := strings.ToLower(status)

				if statusNormalizado == "novo" {
					nomeProvedor := row[0]
					telefone := row[3]

					fmt.Printf("Provedor: %s | Telefone: %s\n", nomeProvedor, telefone)
					contadorNovo++
				}
			}
		}
	}

	fmt.Printf("Total de contatos novos: %d\n", contadorNovo)

}

func gerarAbaLigarHoje() {

	f, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		log.Fatal("erro ao abrir a planilha", err)
	}

	index, err := f.NewSheet("Ligar Hoje")

	f.SetActiveSheet(index)

	rows, err := f.GetRows("Vendas")
	if err != nil {
		log.Fatal("erro ao ler aba vendas")
	}

	for i, headers := range rows[0] {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue("Ligar Hoje", cell, headers)
	}

	linhaNovaAba := 2

	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) >= 7 {
			status := strings.TrimSpace(row[6])
			statusNormalizado := strings.ToLower(status)
			if statusNormalizado == "novo" {
				for colIndex, valorCelula := range row {
					cell, _ := excelize.CoordinatesToCellName(colIndex+1, linhaNovaAba)
					f.SetCellValue("Ligar Hoje", cell, valorCelula)
				}
				linhaNovaAba++
			}
		}
	}

	// salva alterações
	err = f.Save()
	if err != nil {
		log.Fatal("erro ao salvar planilha", err)
	}
	fmt.Println("aba 'Ligar Hoje' criada com sucesso")

}

func main() {

	gerarAbaLigarHoje()

	// opcao := 2

	// if opcao == 1 {
	// 	criarPlanilha()
	// }
	// if opcao == 2 {
	// 	lerPlanilha()
	// }
}
