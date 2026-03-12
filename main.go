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
	Provedor string
	Cidade   string
	Estado   string
	Telefone string
	Contato  string
	Data     string
	Status   string
}

type Dashboard struct {
	Novos      int
	EmLigacao  int
	Ligados    int
	NaoAtendeu int
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

		if len(row) < 7 {
			continue
		}

		status := strings.TrimSpace(row[6])

		if strings.EqualFold(status, statusBusca) {
			contato := Contato{
				Provedor: strings.TrimSpace(row[0]),
				Cidade:   strings.TrimSpace(row[1]),
				Estado:   strings.TrimSpace(row[2]),
				Telefone: strings.TrimSpace(row[3]),
				Contato:  strings.TrimSpace(row[4]),
				Data:     strings.TrimSpace(row[5]),
				Status:   strings.Title(strings.ToLower(strings.TrimSpace(row[6]))),
			}
			contatos = append(contatos, contato)

		}

	}
	return contatos, nil

}

func atualizarStatus(provedor string) error {

	// abre planilha existente
	f, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		return fmt.Errorf("Erro ao abrir a planilha: %w", err)
	}
	defer f.Close()

	// lê todas as linhas da aba Vendas
	rows, err := f.GetRows("Vendas")
	if err != nil {
		return fmt.Errorf("Erro ao ler linhas da aba Vendas: %w", err)
	}

	// percorrer as linhas
	for i, row := range rows {

		if i == 0 {
			continue
		}

		if len(row) < 7 {
			continue
		}

		nomePlanilha := strings.TrimSpace(row[0])
		nomeBusca := strings.TrimSpace(provedor)

		if strings.EqualFold(nomePlanilha, nomeBusca) {
			cellStatus, _ := excelize.CoordinatesToCellName(7, i+1)
			f.SetCellValue("Vendas", cellStatus, "Ligado")
			break
		}
	}

	err = f.Save()
	if err != nil {
		return fmt.Errorf("erro ao salvar planilha: %w", err)
	}
	return nil
}

func statusClass(status string) string {

	switch status {

	case "Novo":
		return "status-novo"

	case "Em Ligação":
		return "status-em-ligacao"

	case "Ligado":
		return "status-ligado"

	case "Não Atendeu":
		return "status-nao-atendeu"
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
		if len(row) < 7 {
			continue
		}

		nomePlanilha := strings.TrimSpace(row[0])
		nomeBusca := strings.TrimSpace(provedor)

		if strings.EqualFold(nomePlanilha, nomeBusca) {
			cellStatus, _ := excelize.CoordinatesToCellName(7, i+1)
			f.SetCellValue("Vendas", cellStatus, "Em ligação")
			break
		}
	}
	return f.Save()

}

func atualizarStatusNaoAtendeu(provedor string) error {
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

		if len(row) < 7 {
			continue
		}

		nomePlanilha := strings.TrimSpace(row[0])
		nomeBusca := strings.TrimSpace(provedor)

		if strings.EqualFold(nomePlanilha, nomeBusca) {
			cellStatus, _ := excelize.CoordinatesToCellName(7, i+1)
			f.SetCellValue("Vendas", cellStatus, "Não Atendeu")
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

		if len(row) < 7 {
			continue
		}

		status := strings.TrimSpace(row[6])

		switch strings.ToLower(status) {
		case "novo":
			dash.Novos++
		case "em ligação":
			dash.EmLigacao++
		case "ligado":
			dash.Ligados++
		case "não atendeu":
			dash.NaoAtendeu++
		}
	}
	return dash, nil
}

func main() {

	r := gin.Default()

	r.SetFuncMap(template.FuncMap{
		"statusClass": statusClass,
	})

	// carrega HTML
	r.LoadHTMLGlob("templates/*")

	r.GET("/", func(c *gin.Context) {
		dash, _ := GerarDashboard()

		c.HTML(http.StatusOK, "index.html", gin.H{
			"dashboard": dash,
		})
	})

	// rota principal
	// r.GET("/", func(c *gin.Context) {
	// 	c.HTML(http.StatusOK, "index.html", nil)
	// })

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
			"message": "Planilha gerada com sucesso!",
		})

	})

	// rota listar novos
	r.GET("/listar", func(c *gin.Context) {

		contatos, err := lerPlanilha()
		dash, _ := GerarDashboard()

		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message":   "Feche a planilha antes de gerar a aba.",
				"dashboard": dash,
			})
			return
		}

		c.HTML(http.StatusOK, "index.html", gin.H{
			"contatos": contatos,
		})
	})

	r.POST("/liguei", func(c *gin.Context) {

		provedor := c.PostForm("provedor")
		dash, _ := GerarDashboard()

		err := atualizarStatus(provedor)
		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message":   "erro ao atualizar status",
				"dashboard": dash,
			})
			return
		}

		contatos, _ := lerPlanilha()

		c.HTML(http.StatusOK, "index.html", gin.H{
			"message":  "status atualizado para ligado",
			"contatos": contatos,
		})
	})

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

	r.POST("/nao-atendeu", func(c *gin.Context) {
		provedor := c.PostForm("provedor")
		dash, _ := GerarDashboard()

		err := atualizarStatusNaoAtendeu(provedor)
		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message": "erro ao atualizar status",
			})
			return
		}
		contatos, _ := lerPlanilha()

		c.HTML(http.StatusOK, "index.html", gin.H{
			"message":   "status atualizado para Não Atendeu",
			"contatos":  contatos,
			"dashboard": dash,
		})
	})

	r.POST("/em-ligacao", func(c *gin.Context) {

		provedor := c.PostForm("provedor")

		AtualizarStatusEmLigacao(provedor)

		contatos, _ := lerPlanilha()
		dash, _ := GerarDashboard()

		c.HTML(http.StatusOK, "index.html", gin.H{
			"message":   "Status atualizado para Em ligação",
			"contatos":  contatos,
			"dashboard": dash,
		})
	})

	r.Run(":8080")

}
