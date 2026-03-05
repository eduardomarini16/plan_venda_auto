package main

import (
	"fmt"
	"net/http"
	"strings"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
)

type Contato struct {
	Provedor string
	Cidade   string
	Estado   string
	Telefone string
	Contato  string
	Data     string
	Status   string
}

func criarPlanilha() error {

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
		return fmt.Errorf("Erro ao criar estilo: %w", err)
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
		return fmt.Errorf("Erro ao salvar arquivo: %w", err)
	}

	fmt.Println("Planilha criada com sucesso!!!")
	return nil

}

func lerPlanilha() ([]Contato, error) {

	contatos := []Contato{}

	// abre planilha existente
	f, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		return contatos, fmt.Errorf("Erro ao abrir a planilha: %w", err)
	}

	// lê todas as linhas da aba Vendas
	rows, err := f.GetRows("Vendas")
	if err != nil {
		return contatos, fmt.Errorf("Erro ao ler linhas da aba Vendas: %w", err)
	}

	fmt.Println("Contatos com status NOVO:")
	fmt.Println("-------------------------")

	contadorNovo := 0

	// percorre as linhas ignorando o cabeçalho
	for i, row := range rows {

		if i == 0 {
			continue // pula cabeçalho
		}

		if len(row) >= 7 {
			contato := Contato{
				Provedor: row[0],
				Cidade:   row[1],
				Estado:   row[2],
				Telefone: row[3],
				Contato:  row[4],
				Data:     row[5],
				Status:   row[6],
			}
			status := strings.ToLower(strings.TrimSpace(contato.Status))

			if status == "novo" {
				contatos = append(contatos, contato)
				contadorNovo++
			}
		}
	}
	fmt.Printf("Total de contatos novos: %d\n", contadorNovo)
	return contatos, nil
}

func gerarAbaLigarHoje() error {

	f, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		return fmt.Errorf("erro ao abrir planilha: %w", err)
	}

	nomeAba := "Ligar Hoje"
	indexExistente, err := f.GetSheetIndex(nomeAba)
	if err != nil {
		return fmt.Errorf("erro ao verificar aba: %w", err)

	}

	//Verifica se já existe
	if indexExistente != -1 {
		err = f.DeleteSheet(nomeAba)
		if err != nil {
			return fmt.Errorf("Erro ao deletar aba antiga: %w", err)
		}
		fmt.Println("Aba antiga removida.")
	}

	// cria aba nova limpa
	index, err := f.NewSheet(nomeAba)
	if err != nil {
		return fmt.Errorf("Erro ao criar nova aba: %w", err)
	}

	f.SetActiveSheet(index)

	rows, err := f.GetRows("Vendas")
	if err != nil {
		return fmt.Errorf("Erro ao ler aba Vendas: %w", err)
	}

	// copia cabeçalho
	for i, header := range rows[0] {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(nomeAba, cell, header)
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
				// copia para aba nova
				// copia para aba nova
				for colIndex, valorCelula := range row {
					cell, _ := excelize.CoordinatesToCellName(colIndex+1, linhaNovaAba)
					f.SetCellValue(nomeAba, cell, valorCelula)
				}

				linhaNovaAba++

				// atualiza status na aba original
				cellStatus, _ := excelize.CoordinatesToCellName(7, i+1)
				f.SetCellValue("Vendas", cellStatus, "Em ligação")
			}
		}
	}

	err = f.Save()
	if err != nil {
		return fmt.Errorf("Erro ao salvar planilha: %w", err)
	}

	fmt.Println("Aba 'Ligar Hoje' atualizada com sucesso!")
	return nil
}

func main() {

	r := gin.Default()

	// carrega HTML
	r.LoadHTMLGlob("templates/*")

	// rota principal
	r.GET("/", func(c *gin.Context) {
		c.HTML(http.StatusOK, "index.html", nil)
	})

	// rota criar planilha
	r.POST("/criar", func(c *gin.Context) {

		err := criarPlanilha()

		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message": "Feche a planilha antes de gerar a aba.",
			})
			return

		}
		c.HTML(http.StatusOK, "index.html", gin.H{
			"message": "Aba 'Ligar Hoje' gerada com sucesso!",
		})

	})

	// rota listar novos
	r.GET("/listar", func(c *gin.Context) {

		contatos, err := lerPlanilha()

		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message": "Feche a planilha antes de gerar a aba.",
			})
			return
		}

		c.HTML(http.StatusOK, "index.html", gin.H{
			"contatos": contatos,
		})
	})

	r.POST("/gerar", func(c *gin.Context) {

		err := gerarAbaLigarHoje()

		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message": "Feche a planilha antes de gerar a aba.",
			})
			return
		}

		c.HTML(http.StatusOK, "index.html", gin.H{
			"message": "Aba 'Ligar Hoje' gerada com sucesso!",
		})
	})

	r.Run(":8080")

}
