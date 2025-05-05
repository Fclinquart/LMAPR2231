import pandas as pd

import matplotlib.pyplot as plt
import numpy as np
from mpl_toolkits.axes_grid1.inset_locator import inset_axes
from mpl_toolkits.axes_grid1.inset_locator import mark_inset
plt.rcParams.update({
        # Font and labels
        "font.size": 8,
        "font.family": "sans-serif",
        "axes.labelsize": 8,
        "xtick.labelsize": 8,
        "ytick.labelsize": 8,

        # Lines and markers
        "lines.linewidth": 0.75,
        "lines.markerfacecolor": 'none',
        "lines.markeredgewidth": 0.75,
        "lines.color": "black",

        # Axes
        "axes.edgecolor": "black",
        "axes.linewidth": 0.75,
        "axes.grid": False,  # no grid
        "axes.facecolor": "white",

        # Ticks
        "xtick.color": "black",
        "ytick.color": "black",

        # Legend
        "legend.fontsize": 8,
        "legend.edgecolor": "none",  # No legend border
        "legend.frameon": False,

        # Savefig
        "savefig.facecolor": "white",
        "savefig.edgecolor": "white"
    })


def extract_all_sheets(file_path):
    excel_file = pd.ExcelFile(file_path)
    return [excel_file.parse(sheet) for sheet in excel_file.sheet_names]

def exp1(df):
    # return R, V, I  : exp 1 et 3
    R = df.iloc[:, 0]
    V = df.iloc[:, 1]
    I = df.iloc[:, 2]
    return R, V, I

def exp2(df):
    """
    return t, V, I, Volume for a fixed resitance of 33 Ohm
    """
    t = df.iloc[:, 0]
    V = df.iloc[:, 1]
    I = df.iloc[:, 2]
    Volume = df.iloc[:, 3] # cm^3
    return t, V, I, Volume


def experiment1_I_vs_V_vs_R(V, I, R):
    fig, ax1 = plt.subplots(figsize=(6, 4))  # Set consistent figure size

    # First y-axis
    ax1.plot(I, V, label="Current vs Voltage", color="red", marker="o", markersize=8, linewidth=0.75, markerfacecolor='white')
    ax1.set_xlabel("Current (A)", fontsize=10)
    ax1.set_ylabel("Voltage (V)", fontsize=10)

    # Second y-axis
    ax2 = ax1.twinx()
    ax2.plot(I, R, label="Current vs Resistance", color="Black", marker="s", markersize=8, linewidth=0.75, markerfacecolor='white')
    ax2.set_ylabel("Resistance (Ohms)", fontsize=10)

    # Combine legend
    lines = ax1.get_lines() + ax2.get_lines()
    labels = [line.get_label() for line in lines]
    ax1.legend(lines, labels, loc="upper left", fontsize=8)

    fig.tight_layout()
    plt.savefig("Output/Exp1.pdf", dpi=300, bbox_inches='tight')

def experiment3_I_vs_V_vs_R(V, I, R):
    fig, ax1 = plt.subplots(figsize=(6, 4))  # Set consistent figure size

    # First y-axis
    ax1.plot(I, V, label="Current vs Voltage", color="red", marker="o", markersize=8, linewidth=0.75, markerfacecolor='white')
    ax1.set_xlabel("Current (mA)", fontsize=10)
    ax1.set_ylabel("Voltage (V)", fontsize=10)

    # Second y-axis
    ax2 = ax1.twinx()
    ax2.plot(I, R, label="Current vs Resistance", color="Black", marker="s", markersize=8, linewidth=0.75, markerfacecolor='white')
    ax2.set_ylabel("Resistance (Ohms)", fontsize=10)

    # Combine legend
    lines = ax1.get_lines() + ax2.get_lines()
    labels = [line.get_label() for line in lines]
    ax1.legend(lines, labels, loc="upper left", fontsize=8)

    fig.tight_layout()
    plt.savefig("Output/Exp3.pdf", dpi=300, bbox_inches='tight')

def experiment3_I_vs_V_P(V, I):
    fig, ax1 = plt.subplots(figsize=(6, 4))  # Set consistent figure size

    # First y-axis
    ax1.plot(I, V, label="Current vs Voltage", color="red", marker="o", markersize=8, linewidth=0.75, markerfacecolor='white')
    ax1.set_xlabel("Current (mA)", fontsize=10)
    ax1.set_ylabel("Voltage (V)", fontsize=10)

    # Second y-axis
    ax2 = ax1.twinx()
    P = V * I
    ax2.plot(I, P * 1e3, label="Current vs Power", color="Black", marker="s", markersize=8, linewidth=0.75, markerfacecolor='white')
    ax2.set_ylabel("Power (mW)", fontsize=10)

    # Combine legend
    lines = ax1.get_lines() + ax2.get_lines()
    labels = [line.get_label() for line in lines]
    ax1.legend(lines, labels, loc="upper left", fontsize=8)

    fig.tight_layout()
    plt.savefig("Output/Exp3_power.pdf", dpi=300, bbox_inches='tight')

def experiment3_R_P(V, I, R):
    fig, ax1 = plt.subplots(figsize=(6, 4))  # Set consistent figure size

    # Main plot
    ax1.plot(R, V * I, label="Resistance vs Power", color="black", marker="o", markersize=8, linewidth=0.75, markerfacecolor='white')
    ax1.set_xlabel("Resistance ($\Omega$)", fontsize=10)
    ax1.set_ylabel("Power (W)", fontsize=10)
    plt.legend(loc="upper left", fontsize=8)

    # Zoomed inset
    ax_inset = inset_axes(ax1, width="40%", height="40%", loc="center right", bbox_transform=ax1.transAxes)
    ax_inset.plot(R, V * I, color="black", marker="o", markersize=4, linewidth=0.75, markerfacecolor='white')
    ax_inset.set_xlim(0, 50)  # Zoom in on x-axis from 0 to 50
    ax_inset.set_ylim(0, 4)   # Zoom in on y-axis from 0 to 4
    ax_inset.tick_params(axis='y', which='major', labelsize=6)

    # Connect the zoomed inset to the main plot with black arrows
    mark_inset(ax1, ax_inset, loc1=2, loc2=4, fc="none", ec="black", lw=0.75, linestyle="--")
    

    fig.tight_layout()
    plt.savefig("Output/Exp1_R_vs_power_with_zoom.pdf", dpi=300, bbox_inches='tight')

def V_vs_I(V, I, filename):
    fig, ax1 = plt.subplots(figsize=(6, 4))  # Set consistent figure size

    # First y-axis
    ax1.plot(I, V, label="Current vs Voltage", color="Black", marker="^", markersize=8, linewidth=0.75, markerfacecolor='white')
    ax1.set_xlabel("Current (A)", fontsize=10)
    ax1.set_ylabel("Voltage (V)", fontsize=10)

    fig.tight_layout()
    plt.savefig('Output/{}.pdf'.format(filename), dpi=300, bbox_inches='tight')

def time_vs_volume(t, Volume, theoretical_volume=False):
    fig, ax1 = plt.subplots(figsize=(6, 4))  # Set consistent figure size

    # First y-axis
    ax1.plot(t, Volume, label="Time vs Volume", color="black", marker="D", markersize=8, linewidth=0.75, markerfacecolor='white')
    ax1.set_xlabel("Time (s)", fontsize=10)
    ax1.set_ylabel("Volume ($cm^3$)", fontsize=10)
    
    if theoretical_volume:
        theoretical_volume = theortitical_volume(t, I2)
        ax1.plot(t, theoretical_volume, label="Theoretical Volume", color="red", marker="o", markersize=8, linewidth=0.75, markerfacecolor='white')
        
        # Second y-axis for Faradic efficiency
        ax2 = ax1.twinx()
        faradic_efficiency = Volume / theoretical_volume
        ax2.plot(t, faradic_efficiency, label="Faradic Efficiency", color="blue", marker="s", markersize=8, linewidth=0.75, markerfacecolor='white')
        ax2.set_ylabel("Faradic Efficiency", fontsize=10)
        # Combine legends
        lines = ax1.get_lines() + ax2.get_lines()
        labels = [line.get_label() for line in lines]
        legend = ax1.legend(lines, labels, loc="upper left", fontsize=8, frameon=True)
       
        legend.get_frame().set_linewidth(0.75)    # Set legend border thickness

    fig.tight_layout()
    plt.savefig("Output/Exp4.pdf", dpi=300, bbox_inches='tight')

# extraction

filname = "Fuel cell data.xlsx"

sheets = extract_all_sheets(filname)

R1, V1, I1 = exp1(sheets[0])
R3, V3, I3 = exp1(sheets[2])


t2, V2, I2, Volume2 = exp2(sheets[1])
t4, V4, I4, Volume4 = exp2(sheets[3])


def theortitical_volume(t,i):
    R = 8.314
    T = 298.15
    F = 96485
    P = 101325
    z = 1
    V = (R * T * i*t) / (z*F * P) # m^3
    V = V * 1e6 # cm^3
    return np.array(V)

v_2_theo = theortitical_volume(t2, I2)
print(v_2_theo)

time_vs_volume(t4, Volume4, False)

