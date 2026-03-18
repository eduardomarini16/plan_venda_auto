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

func BuscarContato(termo string) ([]Contato, error) {
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

		provedor := strings.ToLower(strings.TrimSpace(row[0]))
		cidade := strings.ToLower(strings.TrimSpace(row[1]))
		produto := strings.ToLower(strings.TrimSpace(row[6]))

		if strings.Contains(provedor, termo) || strings.Contains(cidade, termo) || strings.Contains(produto, termo) {

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

func SalvarContato(contato Contato) error {

	if strings.TrimSpace(contato.Provedor) == "" {
		return fmt.Errorf("provedor não pode ser vazio")
	}

	if strings.TrimSpace(contato.Telefone) == "" {
		return fmt.Errorf("telefone é obrigatório")
	}

	f, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		return err
	}
	defer f.Close()

	rows, err := f.GetRows("Vendas")
	if err != nil {
		return err
	}

	novaLinha := len(rows) + 1
	if novaLinha < 2 {
		novaLinha = 2
	}
	f.SetCellValue("Vendas", fmt.Sprintf("A%d", novaLinha), contato.Provedor)
	f.SetCellValue("Vendas", fmt.Sprintf("B%d", novaLinha), contato.Cidade)
	f.SetCellValue("Vendas", fmt.Sprintf("C%d", novaLinha), contato.Estado)
	f.SetCellValue("Vendas", fmt.Sprintf("D%d", novaLinha), contato.Telefone)
	f.SetCellValue("Vendas", fmt.Sprintf("E%d", novaLinha), contato.Contato)
	f.SetCellValue("Vendas", fmt.Sprintf("F%d", novaLinha), contato.Data)
	f.SetCellValue("Vendas", fmt.Sprintf("G%d", novaLinha), contato.Produto)
	f.SetCellValue("Vendas", fmt.Sprintf("H%d", novaLinha), contato.Status)
	f.SetCellValue("Vendas", fmt.Sprintf("I%d", novaLinha), contato.Observacao)

	return f.Save()
}

func StatusSelectClass(status string) string {
	switch strings.ToLower(status) {
	case "novo":
		return "select-novo"
	case "em contato":
		return "select-contato"
	case "orçamento":
		return "select-orcamento"
	case "negociação":
		return "select-negociacao"
	case "cliente":
		return "select-cliente"
	case "não interessado":
		return "select-nao-interessado"
	default:
		return ""
	}
}

func ListarTodos() ([]Contato, error) {

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

	return contatos, nil
}

func DeletarContato(provedor string) error {

	contatos, err := ListarTodos()
	if err != nil {
		return err
	}

	var novos []Contato

	for _, c := range contatos {
		if c.Provedor != provedor {
			novos = append(novos, c)
		}
	}

	return SalvarTodos(novos)
}

func SalvarTodos(contatos []Contato) error {

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

func main() {

	r := gin.Default()

	r.SetFuncMap(template.FuncMap{
		"statusClass": statusClass,
	})

	r.SetFuncMap(template.FuncMap{
		"StatusSelectClass": StatusSelectClass,
		"statusClass":       statusClass,
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

		msg := c.Query("msg")
		provedorEditado := c.Query("provedor")

		var message string
		var messageType string

		if msg == "editado" {
			message = "Contato editado com sucesso"
		}

		if msg == "deletado" {
			message = "Contato deletado com sucesso"
			messageType = "message-sucess"
		}

		if msg == "errodelete" {
			message = "Erro ao deletar contato"
			messageType = "message-error"
		}

		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message":   "Erro ao ler planilha",
				"dashboard": dash,
			})
			return
		}

		c.HTML(http.StatusOK, "index.html", gin.H{
			"contatos":        contatos,
			"dashboard":       dash,
			"message":         message,
			"provedorEditado": provedorEditado,
			"messageType":     messageType,
		})
	})

	r.GET("/novo", func(c *gin.Context) {
		dash, _ := GerarDashboard()

		c.HTML(http.StatusOK, "novo.html", gin.H{
			"dashboard": dash,
		})
	})

	r.POST("/novo", func(c *gin.Context) {

		contato := Contato{
			Provedor:   c.PostForm("provedor"),
			Cidade:     c.PostForm("cidade"),
			Estado:     c.PostForm("estado"),
			Telefone:   c.PostForm("telefone"),
			Contato:    c.PostForm("contato"),
			Data:       c.PostForm("data"),
			Produto:    c.PostForm("produto"),
			Status:     "Novo",
			Observacao: c.PostForm("observacao"),
		}

		err := SalvarContato(contato)
		dash, _ := GerarDashboard()

		if err != nil {
			fmt.Println("Erro: ", err)
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message":   "Erro ao salvar contato",
				"dashboard": dash,
			})
			return
		}
		c.HTML(http.StatusOK, "index.html", gin.H{
			"message":   "Contato adiciona com sucesso",
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

	r.GET("/buscar", func(c *gin.Context) {
		termo := strings.ToLower(c.Query("q"))

		contatos, _ := BuscarContato(termo)
		dash, _ := GerarDashboard()

		c.HTML(http.StatusOK, "index.html", gin.H{
			"contatos":  contatos,
			"dashboard": dash,
		})
	})

	//ROTA ABRIR EDIÇÃO
	r.GET("/editar", func(c *gin.Context) {
		provedor := c.Query("provedor")

		contatos, _ := ListarTodos()

		for _, contato := range contatos {
			if strings.EqualFold(strings.TrimSpace(contato.Provedor), strings.TrimSpace(provedor)) {

				c.HTML(200, "editar.html", gin.H{
					"contato": contato,
				})
				return
			}
		}

		c.String(404, "Contato não encontrado")
	})

	// SALVAR EDIÇÃO
	r.POST("/salvar-edicao", func(c *gin.Context) {

		provedorOriginal := c.PostForm("provedor_original")

		file, _ := excelize.OpenFile("controle_vendas_provedores.xlsx")
		defer file.Close()

		rows, _ := file.GetRows("Vendas")

		for i, row := range rows {
			if i == 0 {
				continue
			}

			if len(row) < 9 {
				continue
			}

			if row[0] == provedorOriginal {

				file.SetCellValue("Vendas", fmt.Sprintf("A%d", i+1), c.PostForm("provedor"))
				file.SetCellValue("Vendas", fmt.Sprintf("B%d", i+1), c.PostForm("cidade"))
				file.SetCellValue("Vendas", fmt.Sprintf("C%d", i+1), c.PostForm("estado"))
				file.SetCellValue("Vendas", fmt.Sprintf("D%d", i+1), c.PostForm("telefone"))
				file.SetCellValue("Vendas", fmt.Sprintf("E%d", i+1), c.PostForm("contato"))
				file.SetCellValue("Vendas", fmt.Sprintf("F%d", i+1), "") // data se quiser depois
				file.SetCellValue("Vendas", fmt.Sprintf("G%d", i+1), c.PostForm("produto"))
				file.SetCellValue("Vendas", fmt.Sprintf("H%d", i+1), c.PostForm("status"))
				file.SetCellValue("Vendas", fmt.Sprintf("I%d", i+1), c.PostForm("observacao"))

				break
			}
		}

		file.Save()

		c.Redirect(302, "/listar?msg=editado&provedor="+provedorOriginal)
	})

	r.POST("/deletar", func(c *gin.Context) {

		provedor := c.PostForm("provedor")

		err := DeletarContato(provedor)
		if err != nil {
			c.Redirect(302, "/listar?msg=errodelete")
			return
		}

		c.Redirect(302, "/listar?msg=deletado")
	})

	r.Run(":8080")
}
