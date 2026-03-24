package services

import (
	"fmt"
	"strings"

	"github.com/eduardomarini16/plan_venda_auto/models"
	repository "github.com/eduardomarini16/plan_venda_auto/repository"
	"github.com/xuri/excelize/v2"
)

func ListarPorStatus(statusBusca string) ([]models.Contato, error) {

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

		statusPlanilha := strings.ToLower(strings.TrimSpace(row[7]))
		statusBusca = strings.ToLower(strings.TrimSpace(statusBusca))

		// fmt.Println(statusBusca)
		// fmt.Println(statusPlanilha)

		if strings.Contains(statusPlanilha, statusBusca) {
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

	}
	return contatos, nil

}

func LerPlanilha() ([]models.Contato, error) {

	contatos, err := ListarPorStatus("novo")
	if err != nil {
		return nil, err
	}

	fmt.Println("Contatos com o status 'NOVO: ")
	fmt.Println("------------------------------")
	fmt.Printf("Total de contatos novos: %d\n", len(contatos))

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

func StatusClass(status string) string {

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

func GerarDashboard() (models.Dashboard, error) {

	file, err := excelize.OpenFile("controle_vendas_provedores.xlsx")
	if err != nil {
		return models.Dashboard{}, err
	}
	defer file.Close()

	rows, err := file.GetRows("Vendas")
	if err != nil {
		return models.Dashboard{}, err
	}

	var dash models.Dashboard

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

func BuscarContato(termo string) ([]models.Contato, error) {
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

		provedor := strings.ToLower(strings.TrimSpace(row[0]))
		cidade := strings.ToLower(strings.TrimSpace(row[1]))
		produto := strings.ToLower(strings.TrimSpace(row[6]))

		if strings.Contains(provedor, termo) || strings.Contains(cidade, termo) || strings.Contains(produto, termo) {

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
	}

	return contatos, nil

}

func SalvarContato(contato models.Contato) error {

	if strings.TrimSpace(contato.Provedor) == "" {
		return fmt.Errorf("provedor não pode ser vazio")
	}

	if strings.TrimSpace(contato.Telefone) == "" {
		return fmt.Errorf("telefone é obrigatório")
	}

	existe, err := ContatoExiste(contato.Provedor)
	if err != nil {
		return err
	}

	if existe {
		return fmt.Errorf("Contato já existe")
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

func ContatoExiste(provedor string) (bool, error) {
	file, err := excelize.OpenFile("lcontrole_vendas_provedores.xlsx")
	if err != nil {
		return false, err
	}
	defer file.Close()

	rows, err := file.GetRows("Vendas")
	if err != nil {
		return false, err
	}

	for i, row := range rows {
		if i == 0 {
			continue
		}

		if len(row) < 1 {
			continue
		}

		nome := strings.TrimSpace(row[0])

		if strings.EqualFold(nome, strings.TrimSpace(provedor)) {
			return true, nil
		}
	}
	return false, nil
}

func ListarPaginado(pagina int, limite int) ([]models.Contato, int, error) {

	todos, err := repository.ListarTodos()
	if err != nil {
		return nil, 0, err
	}

	total := len(todos)

	inicio := (pagina - 1) * limite
	fim := inicio + limite

	if inicio > total {
		return []models.Contato{}, total, nil
	}

	if fim > total {
		fim = total
	}

	return todos[inicio:fim], total, nil

}
