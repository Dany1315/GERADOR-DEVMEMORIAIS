"""
Módulo de rastreamento de progresso com tempo estimado e decorrido.
Integrado com Streamlit para exibir barras de progresso inteligentes.
"""

import time
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Callable
import streamlit as st


class ProgressTracker:
    """
    Rastreador de progresso com tempo estimado e decorrido.
    
    Características:
    - Calcula tempo estimado baseado em histórico
    - Exibe tempo decorrido em tempo real
    - Suporta múltiplas etapas de processamento
    - Integrado com Streamlit
    """
    
    def __init__(self, total_etapas: int = 1, nome_processo: str = "Processamento"):
        """
        Inicializa o rastreador de progresso.
        
        Args:
            total_etapas: Número total de etapas do processo
            nome_processo: Nome do processo para exibição
        """
        self.total_etapas = total_etapas
        self.nome_processo = nome_processo
        self.etapa_atual = 0
        self.tempo_inicio = None
        self.tempo_por_etapa = {}
        self.etapas_info = {}
        self.tempos_historicos = []  # Para calcular média
        
    def iniciar(self):
        """Inicia o rastreamento de tempo."""
        self.tempo_inicio = time.time()
        self.etapa_atual = 0
        
    def atualizar_etapa(self, numero_etapa: int, descricao: str = "", tempo_estimado_seg: Optional[float] = None):
        """
        Atualiza a etapa atual do processo.
        
        Args:
            numero_etapa: Número da etapa (1-indexed)
            descricao: Descrição da etapa
            tempo_estimado_seg: Tempo estimado para esta etapa em segundos
        """
        if self.tempo_inicio is None:
            self.iniciar()
            
        self.etapa_atual = numero_etapa
        self.etapas_info[numero_etapa] = {
            'descricao': descricao,
            'tempo_estimado': tempo_estimado_seg,
            'tempo_inicio': time.time()
        }
        
    def finalizar_etapa(self, numero_etapa: int):
        """Finaliza uma etapa e registra o tempo decorrido."""
        if numero_etapa in self.etapas_info:
            tempo_decorrido = time.time() - self.etapas_info[numero_etapa]['tempo_inicio']
            self.tempo_por_etapa[numero_etapa] = tempo_decorrido
            self.tempos_historicos.append(tempo_decorrido)
            
    def obter_tempo_decorrido(self) -> float:
        """Retorna o tempo decorrido em segundos."""
        if self.tempo_inicio is None:
            return 0
        return time.time() - self.tempo_inicio
    
    def obter_tempo_estimado_restante(self) -> float:
        """Calcula o tempo estimado restante baseado no histórico."""
        if not self.tempos_historicos or self.etapa_atual >= self.total_etapas:
            return 0
            
        tempo_medio_por_etapa = sum(self.tempos_historicos) / len(self.tempos_historicos)
        etapas_restantes = self.total_etapas - self.etapa_atual
        return tempo_medio_por_etapa * etapas_restantes
    
    def obter_tempo_total_estimado(self) -> float:
        """Calcula o tempo total estimado para o processo."""
        if not self.tempos_historicos:
            return 0
            
        tempo_medio_por_etapa = sum(self.tempos_historicos) / len(self.tempos_historicos)
        return tempo_medio_por_etapa * self.total_etapas
    
    def obter_percentual_progresso(self) -> float:
        """Retorna o percentual de progresso (0-100)."""
        if self.total_etapas == 0:
            return 0
        return (self.etapa_atual / self.total_etapas) * 100
    
    def formatar_tempo(self, segundos: float) -> str:
        """Formata segundos em formato legível (HH:MM:SS)."""
        if segundos < 0:
            segundos = 0
            
        horas = int(segundos // 3600)
        minutos = int((segundos % 3600) // 60)
        segs = int(segundos % 60)
        
        if horas > 0:
            return f"{horas}h {minutos}m {segs}s"
        elif minutos > 0:
            return f"{minutos}m {segs}s"
        else:
            return f"{segs}s"
    
    def obter_info_progresso(self) -> Dict:
        """Retorna dicionário com informações de progresso."""
        tempo_decorrido = self.obter_tempo_decorrido()
        tempo_estimado_restante = self.obter_tempo_estimado_restante()
        tempo_total_estimado = self.obter_tempo_total_estimado()
        percentual = self.obter_percentual_progresso()
        
        return {
            'etapa_atual': self.etapa_atual,
            'total_etapas': self.total_etapas,
            'percentual': percentual,
            'tempo_decorrido': tempo_decorrido,
            'tempo_decorrido_formatado': self.formatar_tempo(tempo_decorrido),
            'tempo_estimado_restante': tempo_estimado_restante,
            'tempo_estimado_restante_formatado': self.formatar_tempo(tempo_estimado_restante),
            'tempo_total_estimado': tempo_total_estimado,
            'tempo_total_estimado_formatado': self.formatar_tempo(tempo_total_estimado),
        }


class ProgressBarStreamlit:
    """
    Barra de progresso integrada com Streamlit.
    Exibe progresso, tempo decorrido e tempo estimado.
    """
    
    def __init__(self, tracker: ProgressTracker, container=None):
        """
        Inicializa a barra de progresso do Streamlit.
        
        Args:
            tracker: Instância de ProgressTracker
            container: Container do Streamlit (opcional)
        """
        self.tracker = tracker
        self.container = container or st.container()
        self.progress_bar = None
        self.status_text = None
        
    def atualizar(self, etapa: int, descricao: str = ""):
        """Atualiza a barra de progresso."""
        self.tracker.atualizar_etapa(etapa, descricao)
        self._renderizar()
        
    def finalizar_etapa(self, etapa: int):
        """Finaliza uma etapa."""
        self.tracker.finalizar_etapa(etapa)
        self._renderizar()
        
    def _renderizar(self):
        """Renderiza a barra de progresso no Streamlit."""
        info = self.tracker.obter_info_progresso()
        
        with self.container:
            # Barra de progresso
            percentual = info['percentual'] / 100
            st.progress(percentual, text=f"Etapa {info['etapa_atual']}/{info['total_etapas']}")
            
            # Informações de tempo
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric(
                    "⏱️ Tempo Decorrido",
                    info['tempo_decorrido_formatado'],
                    delta=None
                )
            
            with col2:
                st.metric(
                    "⏳ Tempo Estimado",
                    info['tempo_estimado_restante_formatado'],
                    delta="Restante"
                )
            
            with col3:
                st.metric(
                    "🎯 Total Estimado",
                    info['tempo_total_estimado_formatado'],
                    delta=None
                )


class ProgressBarSimples:
    """
    Barra de progresso simples com status.
    Exibe em formato de texto com emojis.
    """
    
    def __init__(self, tracker: ProgressTracker):
        """
        Inicializa a barra de progresso simples.
        
        Args:
            tracker: Instância de ProgressTracker
        """
        self.tracker = tracker
        
    def obter_barra_visual(self) -> str:
        """Retorna representação visual da barra de progresso."""
        info = self.tracker.obter_info_progresso()
        percentual = info['percentual']
        
        # Cria barra com 20 caracteres
        blocos_preenchidos = int(percentual / 5)  # 100% / 20 = 5%
        blocos_vazios = 20 - blocos_preenchidos
        
        barra = "█" * blocos_preenchidos + "░" * blocos_vazios
        return f"[{barra}] {percentual:.0f}%"
    
    def obter_status_completo(self) -> str:
        """Retorna status completo com todas as informações."""
        info = self.tracker.obter_info_progresso()
        barra = self.obter_barra_visual()
        
        status = f"""
{barra}

📊 Progresso: Etapa {info['etapa_atual']}/{info['total_etapas']}
⏱️  Tempo Decorrido: {info['tempo_decorrido_formatado']}
⏳ Tempo Estimado: {info['tempo_estimado_restante_formatado']}
🎯 Total Estimado: {info['tempo_total_estimado_formatado']}
        """
        return status


def criar_progress_tracker_requerimento() -> ProgressTracker:
    """
    Cria um rastreador de progresso pré-configurado para requerimento de cartório.
    
    Etapas:
    1. Preparando documentos
    2. Analisando com IA
    3. Preenchendo modelo
    4. Finalizando
    """
    tracker = ProgressTracker(total_etapas=4, nome_processo="Geração de Requerimento")
    
    # Pré-configurar etapas
    tracker.etapas_info = {
        1: {'descricao': 'Preparando documentos para análise visual', 'tempo_estimado': 2},
        2: {'descricao': 'Analisando documentos com Gemini IA', 'tempo_estimado': 8},
        3: {'descricao': 'Preenchendo modelo de requerimento Word', 'tempo_estimado': 3},
        4: {'descricao': 'Finalizando documento', 'tempo_estimado': 1},
    }
    
    return tracker


def exemplo_uso():
    """Exemplo de uso do rastreador de progresso."""
    import time
    
    # Criar rastreador
    tracker = ProgressTracker(total_etapas=4, nome_processo="Processamento de Exemplo")
    tracker.iniciar()
    
    # Simular etapas
    for etapa in range(1, 5):
        tracker.atualizar_etapa(etapa, f"Processando etapa {etapa}")
        time.sleep(2)  # Simula processamento
        tracker.finalizar_etapa(etapa)
        
        info = tracker.obter_info_progresso()
        print(f"Etapa {etapa}: {info['tempo_decorrido_formatado']} decorrido, "
              f"{info['tempo_estimado_restante_formatado']} estimado")
    
    print("\n✅ Processo concluído!")
    info_final = tracker.obter_info_progresso()
    print(f"Tempo total: {info_final['tempo_decorrido_formatado']}")


if __name__ == "__main__":
    exemplo_uso()
