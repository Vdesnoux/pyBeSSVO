from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure


class MplCanvas(FigureCanvasQTAgg):
    """
    Canvas Matplotlib intégré Qt (PySide6)
    Layout:
        [ ax_plot ]
        [ ax_metrics (optionnel) ]
    """

    def __init__(self, parent=None):

        # Figure principale
        fig = Figure(constrained_layout=False)
        super().__init__(fig)
        self.setParent(parent)
        self.figure = fig

        # Grille 2 zones
        self.gs = self.figure.add_gridspec(
            2, 1,
            height_ratios=[5, 1]   # plot / metrics
        )

        # Axe principal (spectre)
        self.ax = self.figure.add_subplot(self.gs[0])

        # Axe metrics (désactivé visuellement)
        self.ax_metrics = self.figure.add_subplot(self.gs[1])
        self.ax_metrics.axis("off")

        self._metrics_visible = False
        self._update_layout()

    # ----------------------------
    # PLOT MANAGEMENT
    # ----------------------------

    def clear(self):
        """Efface uniquement le graphe principal."""
        self.ax.clear()

    def reset(self):
        """Reset complet (plot + metrics)."""
        self.ax.clear()
        self.ax_metrics.clear()
        self.ax_metrics.axis("off")
        self._metrics_visible = False

    # ----------------------------
    # METRICS PANEL
    # ----------------------------

    def set_metrics(self, texts):
        """
        Affiche les lignes de debug en bas.
        texts: list[str]
        """

        self.ax_metrics.clear()
        self.ax_metrics.axis("off")

        if texts:
            self.ax_metrics.text(
                0.0, 1.0,
                "\n".join(texts),
                transform=self.ax_metrics.transAxes,
                ha="left",
                va="top",
                fontsize=7,
                family="monospace"
            )
            self._metrics_visible = True
        else:
            self._metrics_visible = False

        self._update_layout()

    def clear_metrics(self):
        """Masque la zone metrics."""
        self.set_metrics([])

    # ----------------------------
    # LAYOUT MANAGEMENT
    # ----------------------------

    def _update_layout(self):
        """
        Ajuste dynamiquement la hauteur des zones.
        """

        if self._metrics_visible:
            self.gs.set_height_ratios([5, 1])
        else:
            self.gs.set_height_ratios([1, 0.0001])

        self.figure.tight_layout()
        self.figure.canvas.draw_idle()

    # ----------------------------
    # HELPERS
    # ----------------------------

    def draw_plot(self):
        """Force redraw propre."""
        self.draw()

    def save(self, path, **kwargs):
        """Export image."""
        self.figure.savefig(path, bbox_inches="tight", **kwargs)