package main

import (
	"fmt"
	"net/http"
	"strings"
	"text/template"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"

	"github.com/eduardomarini16/plan_venda_auto/models"

	repository "github.com/eduardomarini16/plan_venda_auto/repository"
	services "github.com/eduardomarini16/plan_venda_auto/services"
)

func main() {

	r := gin.Default()

	r.SetFuncMap(template.FuncMap{
		"statusClass": services.StatusClass,
	})

	r.SetFuncMap(template.FuncMap{
		"StatusSelectClass": services.StatusSelectClass,
		"statusClass":       services.StatusClass,
	})

	r.LoadHTMLGlob("templates/*")

	// HOME
	r.GET("/", func(c *gin.Context) {
		dash, _ := services.GerarDashboard()

		c.HTML(http.StatusOK, "index.html", gin.H{
			"dashboard": dash,
		})
	})

	// CRIAR PLANILHA
	r.POST("/criar", func(c *gin.Context) {

		err := repository.CriarPlanilha()

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

		contatos, err := services.LerPlanilha()
		dash, _ := services.GerarDashboard()

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
		dash, _ := services.GerarDashboard()

		c.HTML(http.StatusOK, "novo.html", gin.H{
			"dashboard": dash,
		})
	})

	r.POST("/novo", func(c *gin.Context) {

		contato := models.Contato{
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

		err := services.SalvarContato(contato)
		dash, _ := services.GerarDashboard()

		if err != nil {
			fmt.Println("Erro: ", err)
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message":   err.Error(),
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

		contatos, err := services.ListarPorStatus(status)
		dash, _ := services.GerarDashboard()

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

		err := services.AtualizarStatusGenerico(provedor, status)
		dash, _ := services.GerarDashboard()

		if err != nil {
			c.HTML(http.StatusOK, "index.html", gin.H{
				"message":   "Erro ao atualizar status",
				"dashboard": dash,
			})
			return
		}

		contatos, _ := services.LerPlanilha()

		c.HTML(http.StatusOK, "index.html", gin.H{
			"message":   "Status atualizado com sucesso",
			"contatos":  contatos,
			"dashboard": dash,
		})
	})

	r.GET("/buscar", func(c *gin.Context) {
		termo := strings.ToLower(c.Query("q"))

		contatos, _ := services.BuscarContato(termo)
		dash, _ := services.GerarDashboard()

		c.HTML(http.StatusOK, "index.html", gin.H{
			"contatos":  contatos,
			"dashboard": dash,
		})
	})

	//ROTA ABRIR EDIÇÃO
	r.GET("/editar", func(c *gin.Context) {
		provedor := c.Query("provedor")

		contatos, _ := repository.ListarTodos()

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

		err := repository.DeletarContato(provedor)
		if err != nil {
			c.Redirect(302, "/listar?msg=errodelete")
			return
		}

		c.Redirect(302, "/listar?msg=deletado")
	})

	r.Run(":8080")
}
