

import sys
import os
import re
import threading

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox,
    QProgressBar, QTabWidget, QScrollArea, QFrame, QTreeWidget,
    QTreeWidgetItem, QDialog, QComboBox, QSizePolicy, QHeaderView,
)
from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QPixmap, QFont, QIcon

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def resource_path(relative_path: str) -> str:
    """Retourne le chemin correct que ce soit en .py ou en .exe (PyInstaller)."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

from modules.generateur import generer_pv_semestre, generer_pv_annuel
from modules.pv_global_licence import generer_pv_global_licence

# ─── Palette — vert sobre administratif ──────────────────────────────────────
PRIMARY      = "#2E7D32"   # Vert administration (principal)
PRIMARY_DARK = "#1B5E20"   # Vert foncé (survol)
PRIMARY_LIGHT= "#E8F5E9"   # Vert très clair (fond actif)
ACCENT       = "#2E7D32"   # même vert pour les actions
ACCENT_DARK  = "#1B5E20"
DANGER       = "#B71C1C"   # Rouge sobre
DANGER_DARK  = "#7F0000"
BG           = "#FFFFFF"   # fond blanc pur
CARD         = "#FFFFFF"
BORDER       = "#CFD8DC"   # gris-bleu très léger
BORDER_SEC   = "#E0E0E0"   # séparateur secondaire
GRIS_TXT     = "#546E7A"   # texte secondaire bleu-gris
NOIR         = "#212121"   # texte principal presque noir

# ─── Feuille de style globale — flat / administratif ─────────────────────────
STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {BG};
    color: {NOIR};
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 10pt;
}}

/* ── Section plate (bordure gauche verte) ── */
QFrame#section {{
    background-color: {BG};
    border: none;
    border-left: 3px solid {PRIMARY};
}}
QWidget#section_body {{
    background-color: {BG};
}}

/* ── Titre de section ── */
QLabel#section_title {{
    color: {PRIMARY_DARK};
    font-weight: bold;
    font-size: 10pt;
    background: transparent;
    padding: 2px 0px;
}}

/* ── Ligne de séparation entre sections ── */
QFrame#divider {{
    background-color: {BORDER_SEC};
    border: none;
    max-height: 1px;
}}

/* ── Bouton principal (vert sobre) ── */
QPushButton {{
    border: 1px solid {PRIMARY};
    border-radius: 3px;
    padding: 6px 16px;
    font-weight: bold;
    font-size: 9pt;
    color: white;
    background-color: {PRIMARY};
}}
QPushButton:hover {{
    background-color: {PRIMARY_DARK};
    border-color: {PRIMARY_DARK};
}}
QPushButton:disabled {{
    background-color: #BDBDBD;
    border-color: #BDBDBD;
    color: #757575;
}}

/* ── Bouton action principale (génération) ── */
QPushButton#btn_accent {{
    background-color: {PRIMARY};
    border: 1px solid {PRIMARY_DARK};
    padding: 7px 20px;
    font-size: 10pt;
}}
QPushButton#btn_accent:hover {{
    background-color: {PRIMARY_DARK};
}}
QPushButton#btn_accent:disabled {{
    background-color: #BDBDBD;
    border-color: #BDBDBD;
    color: #757575;
}}

/* ── Bouton suppression ── */
QPushButton#btn_danger {{
    background-color: white;
    border: 1px solid {DANGER};
    color: {DANGER};
    padding: 6px 10px;
    font-weight: bold;
}}
QPushButton#btn_danger:hover {{
    background-color: {DANGER};
    color: white;
}}

/* ── Bouton neutre (renommer, etc.) ── */
QPushButton#btn_neutral {{
    background-color: white;
    border: 1px solid {BORDER};
    color: {GRIS_TXT};
}}
QPushButton#btn_neutral:hover {{
    background-color: #ECEFF1;
    border-color: #90A4AE;
    color: {NOIR};
}}

/* ── Champs de saisie ── */
QLineEdit {{
    border: none;
    border-bottom: 1px solid {BORDER};
    border-radius: 0px;
    padding: 5px 4px;
    background-color: {BG};
    color: {NOIR};
    font-size: 9pt;
}}
QLineEdit:read-only {{
    background-color: #FAFAFA;
    color: #607D8B;
}}
QLineEdit:focus {{
    border-bottom: 2px solid {PRIMARY};
}}

/* ── ComboBox ── */
QComboBox {{
    border: 1px solid {BORDER};
    border-radius: 3px;
    padding: 5px 28px 5px 8px;
    background-color: {BG};
    color: {NOIR};
    font-size: 9pt;
    min-height: 24px;
}}
QComboBox:focus {{ border: 1px solid {PRIMARY}; }}
QComboBox:hover {{ border: 1px solid #90A4AE; }}
QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 24px;
    border-left: 1px solid {BORDER};
    border-top-right-radius: 3px;
    border-bottom-right-radius: 3px;
    background-color: #F5F5F5;
}}
QComboBox::down-arrow {{
    width: 10px;
    height: 10px;
    image: none;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 6px solid {GRIS_TXT};
}}
QComboBox::down-arrow:hover {{ border-top-color: {PRIMARY}; }}
QComboBox QAbstractItemView {{
    background: {BG};
    border: 1px solid {BORDER};
    selection-background-color: {PRIMARY_LIGHT};
    selection-color: {NOIR};
    padding: 2px;
    min-width: 500px;
    outline: none;
}}
QComboBox QAbstractItemView::item {{
    padding: 5px 8px;
    min-height: 22px;
}}

/* ── Onglets ── */
QTabWidget::pane {{
    border: none;
    border-top: 2px solid {BORDER_SEC};
    background: {BG};
}}
QTabBar::tab {{
    background: {BG};
    color: {GRIS_TXT};
    font-weight: bold;
    font-size: 10pt;
    padding: 8px 20px;
    border: none;
    border-bottom: 3px solid transparent;
    margin-right: 4px;
}}
QTabBar::tab:selected {{
    color: {PRIMARY};
    border-bottom: 3px solid {PRIMARY};
    background: {BG};
}}
QTabBar::tab:hover:!selected {{
    color: {NOIR};
    background: #F5F5F5;
}}

/* ── TreeWidget ── */
QTreeWidget {{
    background: {BG};
    border: 1px solid {BORDER};
    color: {NOIR};
    font-size: 9pt;
    alternate-background-color: #FAFAFA;
    gridline-color: {BORDER_SEC};
}}
QTreeWidget::item {{
    padding: 3px 0px;
}}
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

/* ── Barre de progression ── */
QProgressBar {{
    border: 1px solid {BORDER};
    border-radius: 2px;
    background: #EEEEEE;
    height: 12px;
    text-align: center;
    font-size: 8pt;
    color: transparent;
}}
QProgressBar::chunk {{
    background-color: {PRIMARY};
    border-radius: 1px;
}}

/* ── Labels secondaires ── */
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

/* ── Scrollbar fine ── */
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
QScrollBar::handle:vertical:hover {{
    background: {PRIMARY};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}

/* ── Barre de statut ── */
QLabel#status_bar {{
    color: {GRIS_TXT};
    font-size: 9pt;
    padding: 3px 14px 4px 14px;
    background: #F5F5F5;
    border-top: 1px solid {BORDER_SEC};
}}

/* ── En-tête principal ── */
QWidget#header_widget {{
    background-color: {BG};
    border-bottom: 2px solid {PRIMARY};
}}

/* ── Séparateur ── */
QFrame#separator {{
    background: {PRIMARY};
    border: none;
}}
"""


# ─── Helpers UI ───────────────────────────────────────────────────────────────

def _btn(text, cmd=None, color="primary", parent=None):
    b = QPushButton(text, parent)
    if color == "accent":
        b.setObjectName("btn_accent")
    elif color == "danger":
        b.setObjectName("btn_danger")
    elif color == "neutral":
        b.setObjectName("btn_neutral")
    if cmd:
        b.clicked.connect(cmd)
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    return b


def _carte(parent_layout, titre, build_fn, header_color=None):
    """Section plate : titre vert + bordure gauche verte + separateur bas."""
    wrapper = QWidget()
    wrapper.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    vbox = QVBoxLayout(wrapper)
    vbox.setContentsMargins(0, 0, 0, 0)
    vbox.setSpacing(0)

    # Bloc avec bordure gauche verte
    section = QFrame()
    section.setObjectName("section")
    section_vbox = QVBoxLayout(section)
    section_vbox.setContentsMargins(14, 10, 14, 12)
    section_vbox.setSpacing(8)

    # Titre
    title_lbl = QLabel(titre)
    title_lbl.setObjectName("section_title")
    section_vbox.addWidget(title_lbl)

    # Corps
    body_w = QWidget()
    body_w.setObjectName("section_body")
    body_layout = QVBoxLayout(body_w)
    body_layout.setContentsMargins(0, 4, 0, 0)
    body_layout.setSpacing(6)
    build_fn(body_layout)
    section_vbox.addWidget(body_w)

    vbox.addWidget(section)

    # Separateur fin en bas
    div = QFrame()
    div.setObjectName("divider")
    div.setFixedHeight(1)
    div.setStyleSheet(f"background-color: {BORDER_SEC}; border: none;")
    vbox.addWidget(div)

    parent_layout.addWidget(wrapper)
    return wrapper


def _fichier_row(parent_layout, line_edit):
    """Ligne Entry readonly + Parcourir + Effacer."""
    row = QWidget()
    hl = QHBoxLayout(row)
    hl.setContentsMargins(0, 0, 0, 0)
    hl.setSpacing(6)
    hl.addWidget(line_edit, 1)
    browse = _btn("Parcourir")
    hl.addWidget(browse)
    clear = _btn("Effacer", color="danger")
    hl.addWidget(clear)
    parent_layout.addWidget(row)
    return browse, clear


def _choisir_fichier(line_edit):
    chemin, _ = QFileDialog.getOpenFileName(
        None, "Sélectionner un fichier PV Excel",
        "", "Excel (*.xlsx *.xls)"
    )
    if chemin:
        line_edit.setText(chemin)


# ─── Worker thread ────────────────────────────────────────────────────────────

class Worker(QObject):
    finished = Signal(str)
    error    = Signal(str)

    def __init__(self, fn):
        super().__init__()
        self._fn = fn

    def run(self):
        try:
            out = self._fn()
            self.finished.emit(out)
        except Exception as e:
            import traceback
            self.error.emit(str(e) + "\n\n" + traceback.format_exc())


# ─── ListeModules ─────────────────────────────────────────────────────────────

class ListeModules(QWidget):
    def __init__(self, label_ajouter="Ajouter une fiche", parent=None):
        super().__init__(parent)
        self.label_ajouter = label_ajouter
        self.entrees: list[dict] = []
        self._build()

    def _build(self):
        vbox = QVBoxLayout(self)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(6)

        # Barre d'actions
        bar = QWidget()
        hl = QHBoxLayout(bar)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(8)
        hl.addWidget(_btn(f"{self.label_ajouter}", self._ajouter))
        hl.addWidget(_btn("Renommer", self._renommer, color="neutral"))
        hl.addWidget(_btn("Supprimer", self._supprimer, color="danger"))
        hl.addStretch()
        vbox.addWidget(bar)

        # TreeWidget (tableau)
        self.tree = QTreeWidget()
        self.tree.setColumnCount(2)
        self.tree.setHeaderLabels(["Nom du module", "Fichier"])
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(False)
        self.tree.setSelectionMode(QTreeWidget.SelectionMode.SingleSelection)
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.tree.header().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.tree.setColumnWidth(0, 220)
        self.tree.setMinimumHeight(130)
        vbox.addWidget(self.tree)

    def _ajouter(self):
        chemins, _ = QFileDialog.getOpenFileNames(
            self, "Sélectionner les fiches de notes",
            "", "Excel (*.xlsx *.xls);;Tous (*.*)"
        )
        for chemin in chemins:
            nom = os.path.splitext(os.path.basename(chemin))[0]
            nom = re.sub(r"^\d+_", "", nom)
            self.entrees.append({"nom": nom, "chemin": chemin})
            item = QTreeWidgetItem([nom, os.path.basename(chemin)])
            self.tree.addTopLevelItem(item)

    def _renommer(self):
        items = self.tree.selectedItems()
        if not items:
            QMessageBox.warning(self, "Attention", "Sélectionnez une ligne à renommer.")
            return
        idx    = self.tree.indexOfTopLevelItem(items[0])
        ancien = self.entrees[idx]["nom"]

        dlg = QDialog(self)
        dlg.setWindowTitle("Renommer le module")
        dlg.setFixedSize(360, 120)
        dlg.setStyleSheet(f"background:{BG};")
        vb = QVBoxLayout(dlg)
        vb.addWidget(QLabel("Nom du module :"))
        ent = QLineEdit(ancien)
        ent.selectAll()
        vb.addWidget(ent)
        btn_ok = _btn("OK", parent=dlg)
        vb.addWidget(btn_ok)

        def _ok():
            nouveau = ent.text().strip()
            if nouveau:
                self.entrees[idx]["nom"] = nouveau
                items[0].setText(0, nouveau)
            dlg.accept()

        btn_ok.clicked.connect(_ok)
        ent.returnPressed.connect(_ok)
        dlg.exec()

    def _supprimer(self):
        items = self.tree.selectedItems()
        if not items:
            return
        idx = self.tree.indexOfTopLevelItem(items[0])
        self.entrees.pop(idx)
        self.tree.takeTopLevelItem(idx)

    def get(self) -> list[tuple[str, str]]:
        return [(e["nom"], e["chemin"]) for e in self.entrees]

    def vider(self):
        self.entrees.clear()
        self.tree.clear()


# ─── Chargement des filières ──────────────────────────────────────────────────

def _charger_filieres() -> tuple[str, list[str]]:
    """Retourne (ecole, [noms de filières]) depuis filieres.json."""
    json_path = resource_path("filieres.json")
    if not os.path.exists(json_path):
        return "École Normale Supérieure de Fès", []
    try:
        import json
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        ecole    = data.get("ecole", "École Normale Supérieure de Fès")
        filieres = [fil["nom"] for fil in data.get("filieres", [])]
        return ecole, filieres
    except Exception:
        return "École Normale Supérieure de Fès", []


# ─── PanneauInfo ──────────────────────────────────────────────────────────────

class PanneauInfo(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._ecole_defaut, self._filieres = _charger_filieres()
        self._build()

    def _build(self):
        vbox = QVBoxLayout(self)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(6)

        # Ligne 1 : École
        row1 = QWidget()
        hl1 = QHBoxLayout(row1)
        hl1.setContentsMargins(0, 0, 0, 0)
        lbl_ecole = QLabel("École :")
        lbl_ecole.setObjectName("label_secondary")
        lbl_ecole.setFixedWidth(72)
        lbl_ecole.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.edit_ecole = QLineEdit(self._ecole_defaut)
        hl1.addWidget(lbl_ecole)
        hl1.addWidget(self.edit_ecole)
        vbox.addWidget(row1)

        # Ligne 2 : Filière (combo) + Année
        row2 = QWidget()
        hl2 = QHBoxLayout(row2)
        hl2.setContentsMargins(0, 0, 0, 0)
        lbl_fil = QLabel("Filière :")
        lbl_fil.setObjectName("label_secondary")
        lbl_fil.setFixedWidth(72)
        lbl_fil.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.combo_filiere = QComboBox()
        self.combo_filiere.setEditable(True)          # permet aussi la saisie libre
        self.combo_filiere.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.combo_filiere.addItem("")                # choix vide en tête
        self.combo_filiere.addItems(self._filieres)
        self.combo_filiere.setCurrentIndex(0)
        self.combo_filiere.setMinimumWidth(200)
        self.combo_filiere.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        lbl_ann = QLabel("Année :")
        lbl_ann.setObjectName("label_secondary")
        lbl_ann.setContentsMargins(12, 0, 0, 0)
        self.edit_annee = QLineEdit("2025-2026")
        self.edit_annee.setFixedWidth(110)

        hl2.addWidget(lbl_fil)
        hl2.addWidget(self.combo_filiere, 1)
        hl2.addWidget(lbl_ann)
        hl2.addWidget(self.edit_annee)
        vbox.addWidget(row2)

        # Ligne 3 : Langue
        row3 = QWidget()
        hl3 = QHBoxLayout(row3)
        hl3.setContentsMargins(0, 0, 0, 0)
        lbl_lng = QLabel("Langue :")
        lbl_lng.setObjectName("label_secondary")
        lbl_lng.setFixedWidth(72)
        lbl_lng.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.combo_langue = QComboBox()
        self.combo_langue.addItems(["fr", "ar"])
        self.combo_langue.setFixedWidth(80)
        hl3.addWidget(lbl_lng)
        hl3.addWidget(self.combo_langue)
        hl3.addStretch()
        vbox.addWidget(row3)

    def get(self):
        return {
            "ecole":   self.edit_ecole.text().strip(),
            "filiere": self.combo_filiere.currentText().strip(),
            "annee":   self.edit_annee.text().strip(),
            "langue":  self.combo_langue.currentText(),
        }


# ─── Onglet PV Semestre ───────────────────────────────────────────────────────

class OngletSemestre(QWidget):
    def __init__(self, status_cb, parent=None):
        super().__init__(parent)
        self.status_cb = status_cb
        self._build()

    def _build(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        outer = QVBoxLayout(container)
        outer.setContentsMargins(16, 12, 16, 12)
        outer.setSpacing(10)

        _carte(outer, "Informations générales",       self._build_info)
        _carte(outer, "Label du semestre",             self._build_label)
        _carte(outer, "Fiches de notes du semestre",   self._build_modules, header_color=PRIMARY_DARK)
        _carte(outer, "PV du même semestre — année précédente (optionnel)",
               self._build_pv_sem_prec, header_color="#4A6FA5")
        _carte(outer, "Générer le PV de semestre",     self._build_actions, header_color=ACCENT)
        outer.addStretch()

        scroll.setWidget(container)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(scroll)

    def _build_info(self, p):
        self.info = PanneauInfo()
        p.addWidget(self.info)

    def _build_label(self, p):
        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(10)
        lbl = QLabel("Label semestre :")
        lbl.setObjectName("label_secondary")
        self.combo_label = QComboBox()
        self.combo_label.addItems(["S1", "S2", "S3", "S4", "S5", "S6"])
        self.combo_label.setFixedWidth(90)
        hl.addWidget(lbl)
        hl.addWidget(self.combo_label)
        hl.addStretch()
        p.addWidget(row)

    def _build_modules(self, p):
        self.liste = ListeModules("Ajouter des fiches de notes")
        p.addWidget(self.liste)

    def _build_pv_sem_prec(self, p):
        info_txt = QLabel(
            "Importez le PV du même semestre de l'année précédente (fichier généré par cet outil). "
            "Pour chaque étudiant absent des fiches du semestre courant, "
            "sa note sera récupérée depuis ce PV et copiée dans la feuille module correspondante."
        )
        info_txt.setObjectName("label_hint")
        info_txt.setWordWrap(True)
        p.addWidget(info_txt)
        self.edit_pv_prec = QLineEdit()
        self.edit_pv_prec.setReadOnly(True)
        self.edit_pv_prec.setPlaceholderText("Aucun fichier sélectionné (optionnel)")
        b, c = _fichier_row(p, self.edit_pv_prec)
        b.clicked.connect(lambda: _choisir_fichier(self.edit_pv_prec))
        c.clicked.connect(lambda: self.edit_pv_prec.clear())

    def _build_actions(self, p):
        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(14)
        self.btn_gen = _btn("Générer le PV de semestre", self._generer, color="accent")
        self.prog = QProgressBar()
        self.prog.setRange(0, 0)
        self.prog.setFixedWidth(200)
        self.prog.setVisible(False)
        hl.addWidget(self.btn_gen)
        hl.addWidget(self.prog)
        hl.addStretch()
        p.addWidget(row)

    def _generer(self):
        fichiers = self.liste.get()
        if not fichiers:
            QMessageBox.warning(self, "Attention", "Ajoutez au moins une fiche de notes.")
            return
        dossier = QFileDialog.getExistingDirectory(self, "Dossier de sauvegarde")
        if not dossier:
            return
        info  = self.info.get()
        label = self.combo_label.currentText()
        pv_prec = self.edit_pv_prec.text().strip() or None

        self.btn_gen.setEnabled(False)
        self.prog.setVisible(True)
        self.status_cb("Génération en cours…")

        def task():
            return generer_pv_semestre(
                fichiers_modules=fichiers, label_semestre=label,
                chemin_pv_prec=pv_prec, dossier_sortie=dossier,
                ecole=info["ecole"], filiere=info["filiere"],
                annee=info["annee"], langue=info["langue"],
                chemin_pv_sem_prec=pv_prec,
            )

        self._thread = QThread()
        self._worker = Worker(task)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._fin)
        self._worker.error.connect(self._erreur)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.start()

    def _fin(self, out):
        self.prog.setVisible(False)
        self.btn_gen.setEnabled(True)
        self.status_cb("✔  PV semestre généré.")
        QMessageBox.information(self, "Succès",
            f"✅ Fichier généré :\n\n{os.path.basename(out)}\n\nDossier :\n{os.path.dirname(out)}")

    def _erreur(self, msg):
        self.prog.setVisible(False)
        self.btn_gen.setEnabled(True)
        self.status_cb("❌ Erreur.")
        QMessageBox.critical(self, "Erreur", msg[:800])


# ─── Onglet PV Annuel ─────────────────────────────────────────────────────────

class OngletAnnuel(QWidget):
    def __init__(self, status_cb, parent=None):
        super().__init__(parent)
        self.status_cb = status_cb
        self._build()

    def _build(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        outer = QVBoxLayout(container)
        outer.setContentsMargins(16, 12, 16, 12)
        outer.setSpacing(10)

        _carte(outer, "Informations générales",                           self._build_info)
        _carte(outer, "Labels des semestres",                             self._build_labels)
        _carte(outer, "Fichier PV — Semestre impair (S1, S3, S5…)",      self._build_pv_si, header_color=PRIMARY_DARK)
        _carte(outer, "Fichier PV — Semestre pair (S2, S4, S6…)",        self._build_pv_sp, header_color=PRIMARY_DARK)
        _carte(outer, "PV précédent (optionnel — pour les redoublants)",  self._build_pv_prec, header_color="#4A6FA5")
        _carte(outer, "Générer le PV annuel",                             self._build_actions, header_color=ACCENT)
        outer.addStretch()

        scroll.setWidget(container)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(scroll)

    def _build_info(self, p):
        self.info = PanneauInfo()
        p.addWidget(self.info)

    def _build_labels(self, p):
        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(10)
        lbl_si = QLabel("Semestre impair :")
        lbl_si.setObjectName("label_secondary")
        self.combo_si = QComboBox()
        self.combo_si.addItems(["S1", "S3", "S5"])
        self.combo_si.setFixedWidth(90)
        lbl_sp = QLabel("Semestre pair :")
        lbl_sp.setObjectName("label_secondary")
        lbl_sp.setContentsMargins(16, 0, 0, 0)
        self.combo_sp = QComboBox()
        self.combo_sp.addItems(["S2", "S4", "S6"])
        self.combo_sp.setFixedWidth(90)
        hl.addWidget(lbl_si)
        hl.addWidget(self.combo_si)
        hl.addWidget(lbl_sp)
        hl.addWidget(self.combo_sp)
        hl.addStretch()
        p.addWidget(row)

    def _build_pv_si(self, p):
        hint = QLabel("Sélectionnez le fichier Excel PV du semestre impair généré par cet outil.")
        hint.setObjectName("label_hint")
        p.addWidget(hint)
        self.edit_fichier_si = QLineEdit()
        self.edit_fichier_si.setReadOnly(True)
        self.edit_fichier_si.setPlaceholderText("Sélectionner un fichier…")
        b, c = _fichier_row(p, self.edit_fichier_si)
        b.clicked.connect(lambda: _choisir_fichier(self.edit_fichier_si))
        c.clicked.connect(lambda: self.edit_fichier_si.clear())

    def _build_pv_sp(self, p):
        hint = QLabel("Sélectionnez le fichier Excel PV du semestre pair généré par cet outil.")
        hint.setObjectName("label_hint")
        p.addWidget(hint)
        self.edit_fichier_sp = QLineEdit()
        self.edit_fichier_sp.setReadOnly(True)
        self.edit_fichier_sp.setPlaceholderText("Sélectionner un fichier…")
        b, c = _fichier_row(p, self.edit_fichier_sp)
        b.clicked.connect(lambda: _choisir_fichier(self.edit_fichier_sp))
        c.clicked.connect(lambda: self.edit_fichier_sp.clear())

    def _build_pv_prec(self, p):
        hint = QLabel("Fichier PV de l'année précédente (pour récupérer les notes des redoublants).")
        hint.setObjectName("label_hint")
        p.addWidget(hint)
        self.edit_fichier_prec = QLineEdit()
        self.edit_fichier_prec.setReadOnly(True)
        self.edit_fichier_prec.setPlaceholderText("Aucun fichier sélectionné (optionnel)")
        b, c = _fichier_row(p, self.edit_fichier_prec)
        b.clicked.connect(lambda: _choisir_fichier(self.edit_fichier_prec))
        c.clicked.connect(lambda: self.edit_fichier_prec.clear())

    def _build_actions(self, p):
        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(14)
        self.btn_gen = _btn("Générer le PV annuel", self._generer, color="accent")
        self.prog = QProgressBar()
        self.prog.setRange(0, 0)
        self.prog.setFixedWidth(200)
        self.prog.setVisible(False)
        hl.addWidget(self.btn_gen)
        hl.addWidget(self.prog)
        hl.addStretch()
        p.addWidget(row)

    def _generer(self):
        si   = self.edit_fichier_si.text().strip()
        sp   = self.edit_fichier_sp.text().strip()
        prec = self.edit_fichier_prec.text().strip() or None
        if not si or not sp:
            QMessageBox.warning(self, "Attention",
                "Veuillez sélectionner les deux fichiers PV (semestre impair et pair).")
            return
        dossier = QFileDialog.getExistingDirectory(self, "Dossier de sauvegarde")
        if not dossier:
            return
        info   = self.info.get()
        lbl_si = self.combo_si.currentText()
        lbl_sp = self.combo_sp.currentText()

        self.btn_gen.setEnabled(False)
        self.prog.setVisible(True)
        self.status_cb("Génération en cours…")

        def task():
            return generer_pv_annuel(
                fichier_pv_si=si, fichier_pv_sp=sp,
                label_si=lbl_si, label_sp=lbl_sp,
                chemin_pv_prec=prec, dossier_sortie=dossier,
                ecole=info["ecole"], filiere=info["filiere"],
                annee=info["annee"], langue=info["langue"],
            )

        self._thread = QThread()
        self._worker = Worker(task)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._fin)
        self._worker.error.connect(self._erreur)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.start()

    def _fin(self, out):
        self.prog.setVisible(False)
        self.btn_gen.setEnabled(True)
        self.status_cb("✔  PV annuel généré.")
        QMessageBox.information(self, "Succès",
            f"✅ Fichier généré :\n\n{os.path.basename(out)}\n\nDossier :\n{os.path.dirname(out)}")

    def _erreur(self, msg):
        self.prog.setVisible(False)
        self.btn_gen.setEnabled(True)
        self.status_cb("❌ Erreur.")
        QMessageBox.critical(self, "Erreur", msg[:800])


# ─── Onglet PV Global Licence ─────────────────────────────────────────────────

class OngletLicence(QWidget):
    ANNEES = [
        {"label": "1ère année", "si_def": "S1", "sp_def": "S2", "color": "#1E40AF"},
        {"label": "2ème année", "si_def": "S3", "sp_def": "S4", "color": "#1E4B8E"},
        {"label": "3ème année", "si_def": "S5", "sp_def": "S6", "color": "#1A3A5C"},
    ]

    def __init__(self, status_cb, parent=None):
        super().__init__(parent)
        self.status_cb = status_cb
        self.vars_annees: list[dict] = []
        self._build()

    def _build(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        outer = QVBoxLayout(container)
        outer.setContentsMargins(16, 12, 16, 12)
        outer.setSpacing(10)

        _carte(outer, "Informations générales", self._build_info)

        for i, cfg in enumerate(self.ANNEES):
            combo_si = QComboBox()
            combo_si.addItems(["S1", "S3", "S5"])
            combo_si.setCurrentText(cfg["si_def"])
            combo_si.setFixedWidth(90)

            combo_sp = QComboBox()
            combo_sp.addItems(["S2", "S4", "S6"])
            combo_sp.setCurrentText(cfg["sp_def"])
            combo_sp.setFixedWidth(90)

            v = {
                "fichier":   QLineEdit(),
                "si":        combo_si,
                "sp":        combo_sp,
                "annee_uv":  QLineEdit(),
            }
            v["fichier"].setReadOnly(True)
            v["fichier"].setPlaceholderText("Sélectionner un fichier…" if i == 0 else "Optionnel")
            self.vars_annees.append(v)
            titre = f"PV Annuel — {cfg['label']} (semestres {cfg['si_def']} & {cfg['sp_def']})"
            _carte(outer, titre,
                   lambda p, vv=v, cc=cfg: self._build_annee(p, vv, cc),
                   header_color=cfg["color"])

        _carte(outer, "Générer le PV Global de Licence", self._build_actions, header_color=PRIMARY)
        outer.addStretch()

        scroll.setWidget(container)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(scroll)

    def _build_info(self, p):
        self.info = PanneauInfo()
        p.addWidget(self.info)

    def _build_annee(self, p, vars_i: dict, cfg: dict):
        lbl_txt = QLabel(
            f"Fichier PV annuel {cfg['label']} généré par cet outil "
            f"(contient les semestres {cfg['si_def']} et {cfg['sp_def']})."
        )
        lbl_txt.setObjectName("label_hint")
        p.addWidget(lbl_txt)

        
        b, c = _fichier_row(p, vars_i["fichier"])
        _edit = vars_i["fichier"]
        b.clicked.connect(lambda checked=False, e=_edit: _choisir_fichier(e))
        c.clicked.connect(lambda checked=False, e=_edit: e.clear())
        
        row_l = QWidget()
        hl = QHBoxLayout(row_l)
        hl.setContentsMargins(0, 4, 0, 0)
        hl.setSpacing(8)

        # Label S impair
        lbl_si = QLabel("Label S impair :")
        lbl_si.setObjectName("label_secondary")
        hl.addWidget(lbl_si)
        hl.addWidget(vars_i["si"])

        # Label S pair
        lbl_sp = QLabel("Label S pair :")
        lbl_sp.setObjectName("label_secondary")
        lbl_sp.setContentsMargins(8, 0, 0, 0)
        hl.addWidget(lbl_sp)
        hl.addWidget(vars_i["sp"])

        # Année universitaire
        lbl_av = QLabel("Année universitaire :")
        lbl_av.setObjectName("label_secondary")
        lbl_av.setContentsMargins(8, 0, 0, 0)
        vars_i["annee_uv"].setFixedWidth(110)
        hint = QLabel("(ex: 2023-2024)")
        hint.setObjectName("label_hint")
        hl.addWidget(lbl_av)
        hl.addWidget(vars_i["annee_uv"])
        hl.addWidget(hint)
        hl.addStretch()
        p.addWidget(row_l)

    def _build_actions(self, p):
        info_txt = QLabel(
            "Au moins le PV de 1ère année est requis. "
            "Les 2ème et 3ème années sont optionnelles (les semestres manquants apparaîtront avec «—»)."
        )
        info_txt.setObjectName("label_hint")
        info_txt.setWordWrap(True)
        p.addWidget(info_txt)

        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(14)
        self.btn_gen = _btn("Générer le PV Global de Licence", self._generer)
        self.prog = QProgressBar()
        self.prog.setRange(0, 0)
        self.prog.setFixedWidth(200)
        self.prog.setVisible(False)
        hl.addWidget(self.btn_gen)
        hl.addWidget(self.prog)
        hl.addStretch()
        p.addWidget(row)

    def _generer(self):
        pv_1a = self.vars_annees[0]["fichier"].text().strip()
        if not pv_1a:
            QMessageBox.warning(self, "Attention",
                "Le fichier PV de 1ère année est obligatoire.\n"
                "Les 2ème et 3ème années sont optionnelles.")
            return
        for i, v in enumerate(self.vars_annees):
            f = v["fichier"].text().strip()
            if f and not os.path.exists(f):
                QMessageBox.critical(self, "Fichier introuvable",
                    f"Le fichier PV de la {i+1}ème année est introuvable :\n{f}")
                return

        dossier = QFileDialog.getExistingDirectory(self, "Dossier de sauvegarde")
        if not dossier:
            return

        info = self.info.get()

        def _f(i):   return self.vars_annees[i]["fichier"].text().strip() or None
        def _si(i):  return self.vars_annees[i]["si"].currentText()
        def _sp(i):  return self.vars_annees[i]["sp"].currentText()
        def _av(i):  return self.vars_annees[i]["annee_uv"].text().strip()

        self.btn_gen.setEnabled(False)
        self.prog.setVisible(True)
        self.status_cb("Génération du PV Global Licence en cours…")

        def task():
            return generer_pv_global_licence(
                pv_1a=_f(0),   labels_1a=(_si(0), _sp(0)),  annee_1a=_av(0),
                pv_2a=_f(1),   labels_2a=(_si(1), _sp(1)),  annee_2a=_av(1),
                pv_3a=_f(2),   labels_3a=(_si(2), _sp(2)),  annee_3a=_av(2),
                ecole=info["ecole"],
                filiere=info["filiere"],
                annee_lib=info["annee"],
                langue=info["langue"],
                dossier_sortie=dossier,
            )

        self._thread = QThread()
        self._worker = Worker(task)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._fin)
        self._worker.error.connect(self._erreur)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.start()

    def _fin(self, out):
        self.prog.setVisible(False)
        self.btn_gen.setEnabled(True)
        self.status_cb("✔  PV Global Licence généré.")
        QMessageBox.information(self, "Succès",
            f"✅ PV Global de Licence généré :\n\n{os.path.basename(out)}\n\nDossier :\n{os.path.dirname(out)}")

    def _erreur(self, msg):
        self.prog.setVisible(False)
        self.btn_gen.setEnabled(True)
        self.status_cb("❌ Erreur lors de la génération.")
        QMessageBox.critical(self, "Erreur", msg[:800])


# ─── Fenêtre principale ───────────────────────────────────────────────────────

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PV Generator — ENS Fès")
        self.resize(940, 780)
        self.setMinimumSize(820, 600)
        # Icône de la fenêtre
        icon_path = resource_path(os.path.join("assets", "app_icon.ico"))
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_vbox = QVBoxLayout(central)
        main_vbox.setContentsMargins(0, 0, 0, 0)
        main_vbox.setSpacing(0)

        # ── En-tête sobre ──
        hdr = QWidget()
        hdr.setObjectName("header_widget")
        hdr.setStyleSheet(f"background-color:{BG}; border-bottom: 2px solid {PRIMARY};")
        hdr_hl = QHBoxLayout(hdr)
        hdr_hl.setContentsMargins(20, 10, 20, 10)

        logo_path = resource_path(os.path.join("assets", "logo_ens_fes.png"))
        logo_lbl = QLabel()
        if os.path.exists(logo_path):
            pix = QPixmap(logo_path)
            if not pix.isNull():
                logo_lbl.setPixmap(pix.scaledToHeight(52, Qt.TransformationMode.SmoothTransformation))
        else:
            logo_lbl.setText("ENS Fès")
            logo_lbl.setStyleSheet(f"font-size:15pt; font-weight:bold; color:{PRIMARY}; background:transparent;")
        hdr_hl.addWidget(logo_lbl, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        hdr_hl.addStretch()

        txt_frame = QWidget()
        txt_frame.setStyleSheet("background:transparent;")
        txt_vbox = QVBoxLayout(txt_frame)
        txt_vbox.setContentsMargins(0, 0, 0, 0)
        txt_vbox.setSpacing(2)
        lbl_title = QLabel("PV Académique — ENS Fès")
        lbl_title.setStyleSheet(f"font-size:15pt; font-weight:bold; color:{PRIMARY}; background:transparent; letter-spacing:0.5px;")
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignRight)
        lbl_sub = QLabel("Génération automatique des procès-verbaux de délibération")
        lbl_sub.setStyleSheet(f"font-size:9pt; color:{GRIS_TXT}; background:transparent;")
        lbl_sub.setAlignment(Qt.AlignmentFlag.AlignRight)
        txt_vbox.addWidget(lbl_title)
        txt_vbox.addWidget(lbl_sub)
        hdr_hl.addWidget(txt_frame, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        main_vbox.addWidget(hdr)

        # ── Onglets ──
        tabs = QTabWidget()
        tabs.addTab(OngletSemestre(self._set_status), "  PV Semestre  ")
        tabs.addTab(OngletAnnuel(self._set_status),   "  PV Annuel  ")
        tabs.addTab(OngletLicence(self._set_status),  "  PV Licence  ")
        main_vbox.addWidget(tabs, 1)

        # Barre de statut
        self.status_lbl = QLabel("Prêt.")
        self.status_lbl.setObjectName("status_bar")
        main_vbox.addWidget(self.status_lbl)

    def _set_status(self, msg: str):
        self.status_lbl.setText(msg)


# ─── Lancement ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    # Icône globale (barre des tâches Windows / dock macOS)
    _ico  = resource_path(os.path.join("assets", "app_icon.ico"))
    _png  = resource_path(os.path.join("assets", "app_icon.png"))
    for _p in (_ico, _png):
        if os.path.exists(_p):
            app.setWindowIcon(QIcon(_p))
            break
    window = App()
    window.show()
    sys.exit(app.exec())