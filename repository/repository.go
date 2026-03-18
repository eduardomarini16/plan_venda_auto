package repository

import (
	"fmt"
	"strings"

	"github.com/eduardomarini16/plan_venda_auto/models"
	"github.com/xuri/excelize/v2"
)

func CriarPlanilha() error {

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
		"Produto",
		"Status",
		"Observação",
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
		return fmt.Errorf("Erro ao criar estilo: %w", err)
	}

	f.SetCellStyle(sheet, "A1", "I1", style)

	// ajusta largura das colunas
	f.SetColWidth(sheet, "A", "A", 25)
	f.SetColWidth(sheet, "B", "C", 15)
	f.SetColWidth(sheet, "D", "D", 18)
	f.SetColWidth(sheet, "E", "E", 22)
	f.SetColWidth(sheet, "F", "G", 20)
	f.SetColWidth(sheet, "H", "H", 18)
	f.SetColWidth(sheet, "I", "I", 30)

	// Salva o arquivo
	err = f.SaveAs("controle_vendas_provedores.xlsx")
	if err != nil {
		return fmt.Errorf("Erro ao salvar arquivo: %w", err)
	}

	fmt.Println("Planilha criada com sucesso!!!")
	return nil

}

func ListarTodos() ([]models.Contato, error) {

	file, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		return nil, err
	}
	defer file.Close()

	rows, err := file.GetRows("Vendas")
	if err != nil {
		return nil, err
	}

	var contatos []models.Contato

	for i, row := range rows {
		if i == 0 {
			continue
		}

		if len(row) < 9 {
			continue
		}

		contato := models.Contato{
			Provedor:   strings.TrimSpace(row[0]),
			Cidade:     strings.TrimSpace(row[1]),
			Estado:     strings.TrimSpace(row[2]),
			Telefone:   strings.TrimSpace(row[3]),
			Contato:    strings.TrimSpace(row[4]),
			Data:       strings.TrimSpace(row[5]),
			Produto:    strings.TrimSpace(row[6]),
			Status:     strings.TrimSpace(row[7]),
			Observacao: strings.TrimSpace(row[8]),
		}

		contatos = append(contatos, contato)
	}

	return contatos, nil
}

func SalvarTodos(contatos []models.Contato) error {

	f := excelize.NewFile()
	sheet := "Vendas"
	f.SetSheetName("Sheet1", sheet)

	// Cabeçalho
	headers := []string{
		"Nome do Provedor",
		"Cidade",
		"Estado",
		"Telefone",
		"Nome do Contato",
		"Data do Primeiro Contato",
		"Produto",
		"Status",
		"Observação",
	}

	for i, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheet, cell, header)
	}

	// Dados
	for i, c := range contatos {
		linha := i + 2

		f.SetCellValue(sheet, fmt.Sprintf("A%d", linha), c.Provedor)
		f.SetCellValue(sheet, fmt.Sprintf("B%d", linha), c.Cidade)
		f.SetCellValue(sheet, fmt.Sprintf("C%d", linha), c.Estado)
		f.SetCellValue(sheet, fmt.Sprintf("D%d", linha), c.Telefone)
		f.SetCellValue(sheet, fmt.Sprintf("E%d", linha), c.Contato)
		f.SetCellValue(sheet, fmt.Sprintf("F%d", linha), c.Data)
		f.SetCellValue(sheet, fmt.Sprintf("G%d", linha), c.Produto)
		f.SetCellValue(sheet, fmt.Sprintf("H%d", linha), c.Status)
		f.SetCellValue(sheet, fmt.Sprintf("I%d", linha), c.Observacao)
	}

	return f.SaveAs("controle_vendas_provedores.xlsx")
}

func DeletarContato(provedor string) error {

	contatos, err := ListarTodos()
	if err != nil {
		return err
	}

	var novos []models.Contato

	for _, c := range contatos {
		if c.Provedor != provedor {
			novos = append(novos, c)
		}
	}

	return SalvarTodos(novos)
}
