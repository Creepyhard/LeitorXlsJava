package br.com.leitor.xls;

public class Cliente {
	private String nome;
    private String protocolo;
    private double glosa;
    private double pedido;
    private double valorfinal;
    private boolean aprovado;

    public Cliente() { }

    public Cliente(String nome, String protocolo, double glosa, double pedido,
                 double valorfinal, boolean aprovado) {
          super();
          this.nome = nome;
          this.protocolo = protocolo;
          this.glosa = glosa;
          this.pedido = pedido;
          this.valorfinal = valorfinal;
          this.aprovado = aprovado;
    }



    public String getNome() {
          return nome;
    }
    public void setNome(String nome) {
          this.nome = nome;
    }
    public String getProtocolo() {
          return protocolo;
    }
    public void setProtocolo(String protocolo) {
          this.protocolo = protocolo;
    }
    public double getGlosa() {
          return glosa;
    }
    public void setGlosa(double glosa) {
          this.glosa = glosa;
    }
    public double getPedido() {
          return pedido;
    }
    public void setPedido(double pedido) {
          this.pedido = pedido;
    }
    public double getValorFinal() {
          return valorfinal;
    }
    public void setValorFinal(double valorfinal) {
          this.valorfinal = valorfinal;
    }

    public boolean getAprovado() {
          return aprovado;
    }

    public void setAprovado(boolean aprovado) {
          this.aprovado = aprovado;
    }
}
