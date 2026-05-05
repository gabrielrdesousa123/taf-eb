const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  // Militares
  militares: {
    listar: () => ipcRenderer.invoke('militares:listar'),
    salvar: (dados) => ipcRenderer.invoke('militares:salvar', dados),
    excluir: (id) => ipcRenderer.invoke('militares:excluir', id),
    buscar: (id) => ipcRenderer.invoke('militares:buscar', id),
  },
  // Avaliações
  avaliacoes: {
    listar: () => ipcRenderer.invoke('avaliacoes:listar'),
    salvar: (dados) => ipcRenderer.invoke('avaliacoes:salvar', dados),
    excluir: (id) => ipcRenderer.invoke('avaliacoes:excluir', id),
  },
  // Resultados
  resultados: {
    salvar: (dados) => ipcRenderer.invoke('resultados:salvar', dados),
    listar: (avaliacao_id) => ipcRenderer.invoke('resultados:listar', avaliacao_id),
    listarComMilitares: (avaliacao_id, chamada) => ipcRenderer.invoke('resultados:listarComMilitares', avaliacao_id, chamada),
    porMilitar: (militar_id) => ipcRenderer.invoke('resultados:porMilitar', militar_id),
    buscar: (militar_id, avaliacao_id, chamada) =>
      ipcRenderer.invoke('resultados:buscar', militar_id, avaliacao_id, chamada),
  },
  // DB
  db: {
    backup: () => ipcRenderer.invoke('db:backup'),
    caminho: () => ipcRenderer.invoke('db:caminho'),
  }
});
