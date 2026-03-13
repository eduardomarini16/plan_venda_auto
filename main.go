package main

import (
	"fmt"
	"net/http"
	"strings"
	"text/template"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
)

type Contato struct {
	Provedor   string
	Cidade     string
	Estado     string
	Telefone   string
	Contato    string
	Data       string
	Produto    string
	Status     string
	Observacao string
}

type Dashboard struct {
	Novos          int
	EmContato      int
	Orcamento      int
	Negociacao     int
	Clientes       int
	NaoInteressado int
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

func lerPlanilha() ([]Contato, error) {

	contatos, err := listarPorStatus("novo")
	if err != nil {
		return nil, err
	}

	fmt.Println("Contatos com o status 'NOVO: ")
	fmt.Println("------------------------------")
	fmt.Printf("Total de contatos novos: %d\n", len(contatos))

	return contatos, nil
}

func listarPorStatus(statusBusca string) ([]Contato, error) {

	file, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		return nil, err
	}
	defer file.Close()

	rows, err := file.GetRows("Vendas")
	if err != nil {
		return nil, err
	}

	var contatos []Contato

	for i, row := range rows {
		if i == 0 {
			continue
		}

		if len(row) < 9 {
			continue
		}

		statusPlanilha := strings.ToLower(strings.TrimSpace(row[7]))
		statusBusca = strings.ToLower(strings.TrimSpace(statusBusca))

		// fmt.Println(statusBusca)
		// fmt.Println(statusPlanilha)

		if strings.Contains(statusPlanilha, statusBusca) {
			contato := Contato{
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

	}
	return contatos, nil

}

func AtualizarStatusGenerico(provedor string, novoStatus string) error {
	f, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		return err
	}
	defer f.Close()

	rows, err := f.GetRows("Vendas")
	if err != nil {
		return err
	}

	for i, row := range rows {
		if i == 0 {
			continue
		}

		if len(row) < 9 {
			continue
		}

		nomePlanilha := strings.TrimSpace(row[0])

		if strings.EqualFold(nomePlanilha, provedor) {
			cellStatus, _ := excelize.CoordinatesToCellName(8, i+1)
			f.SetCellValue("Vendas", cellStatus, novoStatus)
			break
		}
	}
	return f.Save()
}

func statusClass(status string) string {

	switch status {

	case "Novo":
		return "status-novo"

	case "Em contato":
		return "status-contato"

	case "Orçamento enviado":
		return "status-orcamento"

	case "Negociação":
		return "status-negociacao"

	case "Cliente":
		return "status-cliente"

	case "Não interessado":
		return "status-nao-interessado"
	}

	return ""
}

func AtualizarStatusEmLigacao(provedor string) error {
	f, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		return err
	}
	defer f.Close()

	rows, err := f.GetRows("Vendas")
	if err != nil {
		return err
	}

	for i, row := range rows {
		if i == 0 {
			continue
		}
		if len(row) < 9 {
			continue
		}

		nomePlanilha := strings.TrimSpace(row[0])
		nomeBusca := strings.TrimSpace(provedor)

		if strings.EqualFold(nomePlanilha, nomeBusca) {
			cellStatus, _ := excelize.CoordinatesToCellName(8, i+1)
			f.SetCellValue("Vendas", cellStatus, "Em ligação")
			break
		}
	}
	return f.Save()

}

func GerarDashboard() (Dashboard, error) {

	file, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		return Dashboard{}, err
	}
	defer file.Close()

	rows, err := file.GetRows("Vendas")
	if err != nil {
		return Dashboard{}, err
	}

	var dash Dashboard

	for i, row := range rows {
		if i == 0 {
			continue
		}

		if len(row) < 9 {
			continue
		}

		status := strings.ToLower(strings.TrimSpace(row[7]))

		switch {
		case strings.Contains(status, "novo"):
			dash.Novos++
		case strings.Contains(status, "contato"):
			dash.EmContato++
		case strings.Contains(status, "orç"):
			dash.Orcamento++
		case strings.Contains(status, "negoc"):
			dash.Negociacao++
		case strings.Contains(status, "client"):
			dash.Clientes++
		case strings.Contains(status, "interess"):
			dash.NaoInteressado++
		}
	}
	return dash, nil
}

func main() {

	r := gin.Default()

	r.SetFuncMap(template.FuncMap{
		"statusClass": statusClass,
	})

	r.LoadHTMLGlob("templates/*")

	// HOME
	r.GET("/", func(c *gin.Context) {
		dash, _ := GerarDashboard()

		c.HTML(http.StatusOK, "index.html", gin.H{
			"dashboard": dash,
		})
	})

	// CRIAR PLANILHA
	r.POST("/criar", func(c *gin.Context) {

		err := criarPlanilha()

		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message": "Feche a planilha antes de gerar a aba.",
			})
			return
		}

		c.HTML(http.StatusOK, "index.html", gin.H{
			"message": "Planilha gerada com sucesso!",
		})
	})

	// LISTAR
	r.GET("/listar", func(c *gin.Context) {

		contatos, err := lerPlanilha()
		dash, _ := GerarDashboard()

		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message":   "Erro ao ler planilha",
				"dashboard": dash,
			})
			return
		}

		c.HTML(http.StatusOK, "index.html", gin.H{
			"contatos":  contatos,
			"dashboard": dash,
		})
	})

	// FILTRAR STATUS
	r.GET("/status", func(c *gin.Context) {

		status := c.Query("status")

		contatos, err := listarPorStatus(status)
		dash, _ := GerarDashboard()

		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message": "Erro ao filtrar contatos",
			})
			return
		}

		c.HTML(http.StatusOK, "index.html", gin.H{
			"contatos":  contatos,
			"dashboard": dash,
		})
	})

	// ATUALIZAR STATUS
	r.POST("/atualizar-status", func(c *gin.Context) {

		provedor := c.PostForm("provedor")
		status := c.PostForm("status")

		err := AtualizarStatusGenerico(provedor, status)
		dash, _ := GerarDashboard()

		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message":   "Erro ao atualizar status",
				"dashboard": dash,
			})
			return
		}

		contatos, _ := lerPlanilha()

		c.HTML(http.StatusOK, "index.html", gin.H{
			"message":   "Status atualizado com sucesso",
			"contatos":  contatos,
			"dashboard": dash,
		})
	})

	r.Run(":8080")
}
