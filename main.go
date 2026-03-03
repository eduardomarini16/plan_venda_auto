package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {

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
