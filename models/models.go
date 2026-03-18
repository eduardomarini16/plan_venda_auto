package models

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
