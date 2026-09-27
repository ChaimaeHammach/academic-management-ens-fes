"""
app.py — Rattrapage Manager (PySide6)
======================================
Même charte graphique que PV Generator :
  - Fond blanc pur, vert sobre administratif
  - Sections plates avec bordure gauche verte
  - Champs à soulignement, boutons carrés
"""

import sys
import os
import threading

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox,
    QProgressBar, QScrollArea, QFrame, QTreeWidget, QTreeWidgetItem,
    QDialog, QComboBox, QSizePolicy, QHeaderView,
)
from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QPixmap, QIcon, QColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core import traiter_fiches, lire_fiche

# ─── Palette — vert sobre administratif (identique PV) ───────────────────────
PRIMARY       = "#2E7D32"
PRIMARY_DARK  = "#1B5E20"
PRIMARY_LIGHT = "#E8F5E9"
DANGER        = "#B71C1C"
DANGER_DARK   = "#7F0000"
WARNING_C     = "#E65100"
BG            = "#FFFFFF"
BORDER        = "#CFD8DC"
BORDER_SEC    = "#E0E0E0"
GRIS_TXT      = "#546E7A"
NOIR          = "#212121"

ANNEES    = ["2023-2024", "2024-2025", "2025-2026", "2026-2027"]
SEMESTRES = ["S1", "S2", "S3", "S4", "S5", "S6"]

# ─── Feuille de style (identique PV) ─────────────────────────────────────────
STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {BG};
    color: {NOIR};
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 10pt;
}}
QFrame#section {{
    background-color: {BG};
    border: none;
    border-left: 3px solid {PRIMARY};
}}
QWidget#section_body {{ background-color: {BG}; }}
QLabel#section_title {{
    color: {PRIMARY_DARK};
    font-weight: bold;
    font-size: 10pt;
    background: transparent;
    padding: 2px 0px;
}}

QPushButton {{
    border: 1px solid {PRIMARY};
    border-radius: 3px;
    padding: 6px 16px;
    font-weight: bold;
    font-size: 9pt;
    color: white;
    background-color: {PRIMARY};
}}
QPushButton:hover {{ background-color: {PRIMARY_DARK}; border-color: {PRIMARY_DARK}; }}
QPushButton:disabled {{ background-color: #BDBDBD; border-color: #BDBDBD; color: #757575; }}

QPushButton#btn_danger {{
    background-color: white;
    border: 1px solid {DANGER};
    color: {DANGER};
}}
QPushButton#btn_danger:hover {{ background-color: {DANGER}; color: white; }}

QPushButton#btn_warning {{
    background-color: white;
    border: 1px solid {WARNING_C};
    color: {WARNING_C};
}}
QPushButton#btn_warning:hover {{ background-color: {WARNING_C}; color: white; }}

QPushButton#btn_neutral {{
    background-color: white;
    border: 1px solid {BORDER};
    color: {GRIS_TXT};
}}
QPushButton#btn_neutral:hover {{ background-color: #ECEFF1; border-color: #90A4AE; color: {NOIR}; }}

QPushButton#btn_accent {{
    background-color: {PRIMARY};
    border: 1px solid {PRIMARY_DARK};
    padding: 7px 20px;
    font-size: 10pt;
}}
QPushButton#btn_accent:hover {{ background-color: {PRIMARY_DARK}; }}
QPushButton#btn_accent:disabled {{ background-color: #BDBDBD; border-color: #BDBDBD; color: #757575; }}

QLineEdit {{
    border: none;
    border-bottom: 1px solid {BORDER};
    border-radius: 0px;
    padding: 5px 4px;
    background-color: {BG};
    color: {NOIR};
    font-size: 9pt;
}}
QLineEdit:focus {{ border-bottom: 2px solid {PRIMARY}; }}

QComboBox {{
    border: none;
    border-bottom: 1px solid {BORDER};
    border-radius: 0px;
    padding: 5px 4px;
    background-color: {BG};
    color: {NOIR};
    font-size: 9pt;
}}
QComboBox:focus {{ border-bottom: 2px solid {PRIMARY}; }}
QComboBox::drop-down {{ border: none; width: 20px; }}
QComboBox QAbstractItemView {{
    background: {BG};
    border: 1px solid {BORDER};
    selection-background-color: {PRIMARY_LIGHT};
    selection-color: {NOIR};
}}

QTreeWidget {{
    background: {BG};
    border: 1px solid {BORDER};
    color: {NOIR};
    font-size: 9pt;
    alternate-background-color: #FAFAFA;
    gridline-color: {BORDER_SEC};
}}
QTreeWidget::item {{ padding: 4px 2px; }}
QTreeWidget::item:selected {{
    background-color: {PRIMARY_LIGHT};
    color: {NOIR};
}}
QHeaderView::section {{
    background-color: #F5F5F5;
    color: {GRIS_TXT};
    font-weight: bold;
    font-size: 9pt;
    padding: 6px 8px;
    border: none;
    border-bottom: 2px solid {PRIMARY};
    border-right: 1px solid {BORDER_SEC};
}}

QProgressBar {{
    border: 1px solid {BORDER};
    border-radius: 2px;
    background: #EEEEEE;
    height: 12px;
    text-align: center;
    font-size: 8pt;
    color: transparent;
}}
QProgressBar::chunk {{ background-color: {PRIMARY}; border-radius: 1px; }}

QLabel#label_hint {{
    color: {GRIS_TXT};
    font-size: 8pt;
    background: transparent;
}}
QLabel#label_secondary {{
    color: {GRIS_TXT};
    font-size: 9pt;
    background: transparent;
}}

QScrollBar:vertical {{
    background: #F5F5F5;
    width: 8px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: #B0BEC5;
    border-radius: 4px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {PRIMARY}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}

QLabel#status_bar {{
    color: {GRIS_TXT};
    font-size: 9pt;
    padding: 3px 14px 4px 14px;
    background: #F5F5F5;
    border-top: 1px solid {BORDER_SEC};
}}
"""


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _btn(text, cmd=None, color="primary", parent=None):
    b = QPushButton(text, parent)
    if color == "danger":
        b.setObjectName("btn_danger")
    elif color == "warning":
        b.setObjectName("btn_warning")
    elif color == "neutral":
        b.setObjectName("btn_neutral")
    elif color == "accent":
        b.setObjectName("btn_accent")
    if cmd:
        b.clicked.connect(cmd)
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    return b


def _section(parent_layout, titre, build_fn):
    """Section plate : bordure gauche verte + séparateur bas."""
    wrapper = QWidget()
    wrapper.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    vbox = QVBoxLayout(wrapper)
    vbox.setContentsMargins(0, 0, 0, 0)
    vbox.setSpacing(0)

    frame = QFrame()
    frame.setObjectName("section")
    fvbox = QVBoxLayout(frame)
    fvbox.setContentsMargins(14, 10, 14, 12)
    fvbox.setSpacing(8)

    title_lbl = QLabel(titre)
    title_lbl.setObjectName("section_title")
    fvbox.addWidget(title_lbl)

    body = QWidget()
    body.setObjectName("section_body")
    body_layout = QVBoxLayout(body)
    body_layout.setContentsMargins(0, 4, 0, 0)
    body_layout.setSpacing(6)
    build_fn(body_layout)
    fvbox.addWidget(body)

    vbox.addWidget(frame)

    div = QFrame()
    div.setFixedHeight(1)
    div.setStyleSheet(f"background-color: {BORDER_SEC}; border: none;")
    vbox.addWidget(div)

    parent_layout.addWidget(wrapper)
    return wrapper


# ─── Worker thread ────────────────────────────────────────────────────────────

class Worker(QObject):
    finished = Signal(object)
    error    = Signal(str)

    def __init__(self, fn):
        super().__init__()
        self._fn = fn

    def run(self):
        try:
            result = self._fn()
            self.finished.emit(result)
        except Exception as e:
            import traceback
            self.error.emit(str(e) + "\n\n" + traceback.format_exc())


# ─── Dialogue édition d'une ligne ────────────────────────────────────────────

class DialogEditer(QDialog):
    def __init__(self, entree: dict, parent=None):
        super().__init__(parent)
        self.entree = entree
        self.setWindowTitle("Modifier le module")
        self.setFixedSize(400, 200)
        self.setStyleSheet(f"background:{BG};")
        self._build()

    def _build(self):
        vbox = QVBoxLayout(self)
        vbox.setContentsMargins(20, 16, 20, 16)
        vbox.setSpacing(10)

        def _row(label_txt, widget):
            row = QWidget()
            hl = QHBoxLayout(row)
            hl.setContentsMargins(0, 0, 0, 0)
            lbl = QLabel(label_txt)
            lbl.setObjectName("label_secondary")
            lbl.setFixedWidth(120)
            lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            hl.addWidget(lbl)
            hl.addWidget(widget, 1)
            vbox.addWidget(row)
            return widget

        self.edit_nom = _row("Nom du module :", QLineEdit(self.entree.get("nom_module") or ""))

        self.combo_annee = QComboBox()
        self.combo_annee.addItems(ANNEES)
        self.combo_annee.setEditable(True)
        cur_annee = self.entree.get("annee") or ""
        if cur_annee in ANNEES:
            self.combo_annee.setCurrentText(cur_annee)
        else:
            self.combo_annee.setCurrentText(cur_annee)
        _row("Année :", self.combo_annee)

        self.combo_sem = QComboBox()
        self.combo_sem.addItems([""] + SEMESTRES)
        self.combo_sem.setCurrentText(self.entree.get("semestre") or "")
        _row("Semestre :", self.combo_sem)

        btn_ok = _btn("Enregistrer", self._ok, color="accent")
        vbox.addWidget(btn_ok, alignment=Qt.AlignmentFlag.AlignRight)

    def _ok(self):
        self.entree["nom_module"] = self.edit_nom.text().strip()
        self.entree["annee"]      = self.combo_annee.currentText().strip()
        self.entree["semestre"]   = self.combo_sem.currentText().strip()
        self.accept()


# ─── Fenêtre principale ───────────────────────────────────────────────────────

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Rattrapage Manager — ENS Fès")
        self.resize(860, 680)
        self.setMinimumSize(720, 520)

        # Icône
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  "assets", "app_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # Données
        self.entrees: list[dict] = []

        self._build_ui()

    # ── Construction UI ───────────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_vbox = QVBoxLayout(central)
        main_vbox.setContentsMargins(0, 0, 0, 0)
        main_vbox.setSpacing(0)

        # En-tête
        hdr = QWidget()
        hdr.setStyleSheet(f"background-color:{BG}; border-bottom: 2px solid {PRIMARY};")
        hdr_hl = QHBoxLayout(hdr)
        hdr_hl.setContentsMargins(20, 10, 20, 10)

        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  "assets", "logo_ens_fes.png")
        logo_lbl = QLabel()
        if os.path.exists(logo_path):
            pix = QPixmap(logo_path)
            if not pix.isNull():
                logo_lbl.setPixmap(pix.scaledToHeight(52, Qt.TransformationMode.SmoothTransformation))
        else:
            logo_lbl.setText("ENS Fès")
            logo_lbl.setStyleSheet(f"font-size:14pt; font-weight:bold; color:{PRIMARY};")
        hdr_hl.addWidget(logo_lbl, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        hdr_hl.addStretch()

        txt = QWidget()
        txt.setStyleSheet("background:transparent;")
        txt_vbox = QVBoxLayout(txt)
        txt_vbox.setContentsMargins(0, 0, 0, 0)
        txt_vbox.setSpacing(2)
        lbl_t = QLabel("Rattrapage Manager — ENS Fès")
        lbl_t.setStyleSheet(f"font-size:15pt; font-weight:bold; color:{PRIMARY}; background:transparent;")
        lbl_t.setAlignment(Qt.AlignmentFlag.AlignRight)
        lbl_s = QLabel("Gestion automatique des listes de rattrapage")
        lbl_s.setStyleSheet(f"font-size:9pt; color:{GRIS_TXT}; background:transparent;")
        lbl_s.setAlignment(Qt.AlignmentFlag.AlignRight)
        txt_vbox.addWidget(lbl_t)
        txt_vbox.addWidget(lbl_s)
        hdr_hl.addWidget(txt, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        main_vbox.addWidget(hdr)

        # Zone scrollable
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        container = QWidget()
        outer = QVBoxLayout(container)
        outer.setContentsMargins(16, 12, 16, 12)
        outer.setSpacing(10)

        _section(outer, "1 — Importer les fiches de notes", self._build_import)
        _section(outer, "2 — Modules chargés",              self._build_liste)
        _section(outer, "3 — Générer les fichiers de rattrapage", self._build_actions)
        outer.addStretch()

        scroll.setWidget(container)
        main_vbox.addWidget(scroll, 1)

        # Barre de statut
        self.status_lbl = QLabel("Prêt.")
        self.status_lbl.setObjectName("status_bar")
        main_vbox.addWidget(self.status_lbl)

    # ── Section 1 : Import ────────────────────────────────────────────────────

    def _build_import(self, p):
        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(10)

        hl.addWidget(_btn("Choisir les fiches", self._importer))
        hl.addWidget(_btn("Tout vider", self._vider_tout, color="danger"))

        self.lbl_nb = QLabel("Aucun fichier sélectionné")
        self.lbl_nb.setObjectName("label_hint")
        hl.addWidget(self.lbl_nb)
        hl.addStretch()
        p.addWidget(row)

    def _importer(self):
        chemins, _ = QFileDialog.getOpenFileNames(
            self, "Sélectionner les fiches de notes",
            "", "Excel (*.xlsx *.xls);;Tous (*.*)"
        )
        if not chemins:
            return
        for c in chemins:
            if not any(e["chemin"] == c for e in self.entrees):
                self.entrees.append({
                    "chemin":     c,
                    "info":       None,
                    "annee":      None,
                    "semestre":   None,
                    "nom_module": None,
                    "erreur":     None,
                })
        self._lire_nouveaux()

    def _lire_nouveaux(self):
        self.status_lbl.setText("Lecture des fiches en cours…")
        for e in self.entrees:
            if e["info"] is not None or e["erreur"] is not None:
                continue
            try:
                info = lire_fiche(e["chemin"])
                e["info"]       = info
                e["annee"]      = info.get("annee") or ""
                e["nom_module"] = info.get("module") or ""
                e["semestre"]   = ""
                e["erreur"]     = None
            except Exception as ex:
                e["erreur"] = str(ex)
        self._actualiser_liste()

    def _vider_tout(self):
        self.entrees = []
        self._actualiser_liste()
        self.status_lbl.setText("Liste vidée.")

    # ── Section 2 : Liste modules ─────────────────────────────────────────────

    def _build_liste(self, p):
        self.tree = QTreeWidget()
        self.tree.setColumnCount(6)
        self.tree.setHeaderLabels(["Nom du module", "Année", "Semestre", "Total", "Rattrapage", "Statut"])
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(False)
        self.tree.setSelectionMode(QTreeWidget.SelectionMode.SingleSelection)
        self.tree.setMinimumHeight(160)
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tree.setColumnWidth(1, 90)
        self.tree.setColumnWidth(2, 80)
        self.tree.setColumnWidth(3, 60)
        self.tree.setColumnWidth(4, 90)
        self.tree.setColumnWidth(5, 120)
        p.addWidget(self.tree)

        # Boutons d'action sur la ligne
        bar = QWidget()
        hl = QHBoxLayout(bar)
        hl.setContentsMargins(0, 4, 0, 0)
        hl.setSpacing(8)
        hl.addWidget(_btn("Modifier la ligne", self._editer_ligne, color="neutral"))
        hl.addWidget(_btn("Reinitialiser",      self._reinitialiser_ligne, color="warning"))
        hl.addWidget(_btn("Supprimer la ligne", self._supprimer_ligne, color="danger"))
        hl.addStretch()
        p.addWidget(bar)

    def _actualiser_liste(self):
        self.tree.clear()
        nb = len(self.entrees)
        if nb == 0:
            self.lbl_nb.setText("Aucun fichier sélectionné")
            self.status_lbl.setText("Prêt.")
            return

        self.lbl_nb.setText(f"{nb} fichier(s) chargé(s)")
        total_ratt = 0

        for e in self.entrees:
            if e["erreur"]:
                vals = [
                    os.path.basename(e["chemin"])[:35],
                    "—", "—", "—", "—",
                    f"Erreur : {e['erreur'][:30]}"
                ]
                item = QTreeWidgetItem(vals)
                item.setBackground(5, QColor("#FFF9C4"))
                item.setForeground(5, QColor("#E65100"))
            else:
                info   = e["info"]
                nb_t   = len(info.get("df_all", []))
                nb_r   = len(info.get("df_ratt", []))
                total_ratt += nb_r
                nom    = e["nom_module"] or info.get("module") or os.path.basename(e["chemin"])
                statut = "OK" if nb_t > 0 else "Fichier vide"
                vals = [
                    nom[:40],
                    e["annee"] or info.get("annee") or "—",
                    e["semestre"] or "—",
                    str(nb_t),
                    str(nb_r),
                    statut,
                ]
                item = QTreeWidgetItem(vals)
                if nb_r > 0:
                    item.setBackground(4, QColor("#FFEBEE"))
                    item.setForeground(4, QColor(DANGER))
                else:
                    item.setBackground(4, QColor("#E8F5E9"))
                    item.setForeground(4, QColor(PRIMARY_DARK))
                for col in range(6):
                    item.setTextAlignment(col, Qt.AlignmentFlag.AlignCenter)
                item.setTextAlignment(0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

            # Stocker référence id pour retrouver l'entrée
            item.setData(0, Qt.ItemDataRole.UserRole, id(e))
            self.tree.addTopLevelItem(item)

        self.status_lbl.setText(
            f"{nb} module(s)  —  {total_ratt} étudiant(s) en rattrapage au total."
        )

    def _get_sel(self):
        items = self.tree.selectedItems()
        if not items:
            QMessageBox.warning(self, "Attention", "Sélectionnez une ligne.")
            return None
        item_id = items[0].data(0, Qt.ItemDataRole.UserRole)
        for e in self.entrees:
            if id(e) == item_id:
                return e
        return None

    def _editer_ligne(self):
        e = self._get_sel()
        if e is None:
            return
        if e["erreur"]:
            QMessageBox.warning(self, "Attention",
                "Ce fichier est en erreur. Utilisez 'Reinitialiser' pour le recharger.")
            return
        dlg = DialogEditer(e, self)
        if dlg.exec():
            self._actualiser_liste()

    def _reinitialiser_ligne(self):
        e = self._get_sel()
        if e is None:
            return
        e["info"] = e["erreur"] = e["annee"] = e["semestre"] = e["nom_module"] = None
        self._lire_nouveaux()
        self.status_lbl.setText(f"{os.path.basename(e['chemin'])} rechargé.")

    def _supprimer_ligne(self):
        e = self._get_sel()
        if e is None:
            return
        self.entrees.remove(e)
        self._actualiser_liste()

    # ── Section 3 : Actions ───────────────────────────────────────────────────

    def _build_actions(self, p):
        desc = QLabel(
            "Fichier 1 — Fiches de rattrapage : même format que l'original, "
            "uniquement les étudiants non validés, une feuille par module.\n"
            "Fichier 2 — Récapitulatif + Matrice des conflits : liste consolidée "
            "+ analyse des conflits d'horaires + planning suggéré."
        )
        desc.setObjectName("label_hint")
        desc.setWordWrap(True)
        p.addWidget(desc)

        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(14)
        self.btn_gen = _btn("Générer les fichiers", self._generer, color="accent")
        self.prog = QProgressBar()
        self.prog.setRange(0, 0)
        self.prog.setFixedWidth(200)
        self.prog.setVisible(False)
        hl.addWidget(self.btn_gen)
        hl.addWidget(self.prog)
        hl.addStretch()
        p.addWidget(row)

    def _generer(self):
        valides = [e for e in self.entrees if e["info"] and not e["erreur"]]
        if not valides:
            QMessageBox.warning(self, "Attention", "Aucune fiche valide chargée.")
            return
        dossier = QFileDialog.getExistingDirectory(self, "Choisir le dossier de sauvegarde")
        if not dossier:
            return

        # Propager les modifications manuelles
        for e in valides:
            if e["nom_module"]:
                e["info"]["module"]   = e["nom_module"]
            if e["annee"]:
                e["info"]["annee"]    = e["annee"]
            if e["semestre"]:
                e["info"]["semestre"] = e["semestre"]

        modules_info = [e["info"] for e in valides]

        self.btn_gen.setEnabled(False)
        self.prog.setVisible(True)
        self.status_lbl.setText("Génération en cours…")

        def task():
            return traiter_fiches(modules_info, dossier)

        self._thread = QThread()
        self._worker = Worker(task)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._fin)
        self._worker.error.connect(self._erreur)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.start()

    def _fin(self, result):
        self.prog.setVisible(False)
        self.btn_gen.setEnabled(True)
        self.status_lbl.setText("Fichiers générés avec succès.")
        f1, f2 = result
        QMessageBox.information(self, "Succès",
            f"Deux fichiers générés :\n\n"
            f"{os.path.basename(f1)}\n"
            f"{os.path.basename(f2)}\n\n"
            f"Dossier :\n{os.path.dirname(f1)}"
        )

    def _erreur(self, msg):
        self.prog.setVisible(False)
        self.btn_gen.setEnabled(True)
        self.status_lbl.setText("Erreur lors de la génération.")
        QMessageBox.critical(self, "Erreur", msg[:900])


# ─── Lancement ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    _base = os.path.dirname(os.path.abspath(__file__))
    for _p in (
        os.path.join(_base, "assets", "app_icon.ico"),
        os.path.join(_base, "assets", "app_icon.png"),
    ):
        if os.path.exists(_p):
            app.setWindowIcon(QIcon(_p))
            break
    window = App()
    window.show()
    sys.exit(app.exec())
