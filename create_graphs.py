import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Lese die Excel-Datei
df = pd.read_excel('F:/Friehe_messungen/Vermessung.xlsx', sheet_name=0, header=1)

# Extrahiere relevante Spalten
freq = pd.to_numeric(df.iloc[:, 1], errors='coerce')  # Frequenz
mittelwert_70x44 = pd.to_numeric(df.iloc[:, 5], errors='coerce')  # Mittelwert 7x4
bose = pd.to_numeric(df.iloc[:, 6], errors='coerce')  # Spaßmessung Bose FRF
mittelwert_1m = pd.to_numeric(df.iloc[:, 10], errors='coerce')  # Mittelwert 10x7 (1m)
mittelwerte_1m_skript = pd.to_numeric(df.iloc[:, 11], errors='coerce')  # Mittelwerte 1m Skript
mittelwerte_70x44_skript = pd.to_numeric(df.iloc[:, 12], errors='coerce')  # Mittelwerte 70x44 Skript

# Entferne NaN-Werte
valid = ~(freq.isna() | mittelwert_70x44.isna() | bose.isna() | mittelwert_1m.isna() | mittelwerte_1m_skript.isna() | mittelwerte_70x44_skript.isna())
freq = freq[valid]
mittelwert_70x44 = mittelwert_70x44[valid]
bose = bose[valid]
mittelwert_1m = mittelwert_1m[valid]
mittelwerte_1m_skript = mittelwerte_1m_skript[valid]
mittelwerte_70x44_skript = mittelwerte_70x44_skript[valid]

# Filter für 20-20000 Hz
mask = (freq >= 20) & (freq <= 20000)
freq = freq[mask]
mittelwert_70x44 = mittelwert_70x44[mask]
bose = bose[mask]
mittelwert_1m = mittelwert_1m[mask]
mittelwerte_1m_skript = mittelwerte_1m_skript[mask]
mittelwerte_70x44_skript = mittelwerte_70x44_skript[mask]

# Stil für wissenschaftliche Graphen
plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.size': 10,
    'axes.labelsize': 11,
    'axes.titlesize': 12,
    'legend.fontsize': 9,
    'xtick.labelsize': 9,
    'ytick.labelsize': 9,
    'figure.figsize': (10, 6),
    'axes.grid': True,
    'grid.alpha': 0.3
})

# Graph 1: Mittelwert 70x44, Mittelwerte 70x44 Skript, Bose
fig1, ax1 = plt.subplots()
ax1.semilogx(freq, mittelwert_70x44, label='Mittelwert 70x44', linewidth=1.5, color='blue')
ax1.semilogx(freq, mittelwerte_70x44_skript, label='Mittelwerte 70x44 Skript', linewidth=1.5, color='green')
ax1.semilogx(freq, bose, label='Bose Lautsprecher', linewidth=1.5, color='red')
ax1.set_xlabel('Frequenz [Hz]')
ax1.set_ylabel('Übertragungsmaß [dB]')
ax1.set_xlim(20, 20000)
ax1.set_ylim(0, 120)
ax1.legend(loc='best')
ax1.set_title('Frequenzgang DML 70x44 cm')
plt.tight_layout()
fig1.savefig('C:/Users/frieh/Desktop/Bachelorarbeit/pictures/frequenzgang_70x44.png', dpi=300, bbox_inches='tight')
plt.close(fig1)

# Graph 2: Mittelwert 1m, Mittelwerte 1m Skript, Bose
fig2, ax2 = plt.subplots()
ax2.semilogx(freq, mittelwert_1m, label='Mittelwert 1m', linewidth=1.5, color='blue')
ax2.semilogx(freq, mittelwerte_1m_skript, label='Mittelwerte 1m Skript', linewidth=1.5, color='green')
ax2.semilogx(freq, bose, label='Bose Lautsprecher', linewidth=1.5, color='red')
ax2.set_xlabel('Frequenz [Hz]')
ax2.set_ylabel('Übertragungsmaß [dB]')
ax2.set_xlim(20, 20000)
ax2.set_ylim(0, 120)
ax2.legend(loc='best')
ax2.set_title('Frequenzgang DML 100x70 cm')
plt.tight_layout()
fig2.savefig('C:/Users/frieh/Desktop/Bachelorarbeit/pictures/frequenzgang_100x70.png', dpi=300, bbox_inches='tight')
plt.close(fig2)

# Graph 3: Alle 5 Messungen
fig3, ax3 = plt.subplots()
ax3.semilogx(freq, mittelwert_70x44, label='Mittelwert 70x44', linewidth=1.5, color='blue')
ax3.semilogx(freq, mittelwerte_70x44_skript, label='Mittelwerte 70x44 Skript', linewidth=1.5, color='cyan')
ax3.semilogx(freq, mittelwert_1m, label='Mittelwert 1m', linewidth=1.5, color='green')
ax3.semilogx(freq, mittelwerte_1m_skript, label='Mittelwerte 1m Skript', linewidth=1.5, color='lime')
ax3.semilogx(freq, bose, label='Bose Lautsprecher', linewidth=1.5, color='red')
ax3.set_xlabel('Frequenz [Hz]')
ax3.set_ylabel('Übertragungsmaß [dB]')
ax3.set_xlim(20, 20000)
ax3.set_ylim(0, 120)
ax3.legend(loc='best')
ax3.set_title('Frequenzgang Vergleich aller Messungen')
plt.tight_layout()
fig3.savefig('C:/Users/frieh/Desktop/Bachelorarbeit/pictures/frequenzgang_vergleich.png', dpi=300, bbox_inches='tight')
plt.close(fig3)

print('Graphen erfolgreich erstellt:')
print('- pictures/frequenzgang_70x44.png (Mittelwert 70x44, 70x44 Skript, Bose)')
print('- pictures/frequenzgang_100x70.png (Mittelwert 1m, 1m Skript, Bose)')
print('- pictures/frequenzgang_vergleich.png (Alle 5 Messungen)')
