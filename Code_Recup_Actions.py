import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime
from cvxopt import matrix, solvers
import matplotlib.pyplot as plt

# Liste des symboles boursiers des actions que vous possédez
symbols = ['TFI.PA', 'TTE', 'BB.PA', 'ML.PA', 'SGO.PA', 'STLA']  # Remplacez par vos actions européennes

# Période de temps (deux dernières années)
end_date = datetime.now()
start_date = end_date - pd.DateOffset(years=2)

# Récupérer les données hebdomadaires
data = {}
for symbol in symbols:
    try:
        stock_data = yf.download(symbol, start=start_date, end=end_date, interval='1wk')
        if not stock_data.empty:
            data[symbol] = stock_data['Adj Close']
        else:
            print(f"No data found for {symbol}")
    except Exception as e:
        print(f"Failed to download {symbol}: {e}")

# Combiner les données en un DataFrame
df = pd.DataFrame(data)

# Supprimer les colonnes avec des valeurs manquantes
df.dropna(axis=1, inplace=True)

if df.empty:
    raise ValueError("No valid data found for the provided symbols.")

# Calculer les rendements hebdomadaires
returns = df.pct_change().dropna()

# Calcul des rendements moyens et de la matrice de covariance
mean_returns = returns.mean()
cov_matrix = returns.cov()

# Convertir en matrices cvxopt
n = len(mean_returns)
P = matrix(cov_matrix.values)
q = matrix(np.zeros(n))
G = matrix(np.vstack((-np.eye(n), np.eye(n))))
h = matrix(np.hstack((np.zeros(n), np.ones(n))))
A = matrix(1.0, (1, n))
b = matrix(1.0)

# Fonction pour calculer le portefeuille efficient pour un rendement cible donné
def efficient_frontier(target_return):
    b2 = matrix([1.0, target_return])
    A2 = matrix(np.vstack((np.ones(n), mean_returns)))
    solvers.options['show_progress'] = False
    sol = solvers.qp(P, q, G, h, A2, b2)
    return np.array(sol['x']).flatten()

# Calcul des portefeuilles efficients pour différents rendements cibles
target_returns = np.linspace(mean_returns.min(), mean_returns.max(), 50)
portfolios = [efficient_frontier(tr) for tr in target_returns]
risks = [np.sqrt(np.dot(p.T, np.dot(cov_matrix, p))) for p in portfolios]
returns = [np.dot(p, mean_returns) for p in portfolios]

# Calcul du portefeuille optimisé pour un rendement cible spécifique
target_return = mean_returns.mean()  # Par exemple, la moyenne des rendements moyens des actions
optimized_weights = efficient_frontier(target_return)
optimized_return = np.dot(optimized_weights, mean_returns)
optimized_risk = np.sqrt(np.dot(optimized_weights.T, np.dot(cov_matrix, optimized_weights)))

# Vérification des résultats
print("Rendement du Portefeuille Optimisé hebdomadaire:", optimized_return)
print("Risque du Portefeuille Optimisé hebdomadaire:", optimized_risk)
annual_return = (1 + optimized_return) ** 52 - 1
print("Rendement Annuel du Porte-feuille(en %):", annual_return*100)
annual_risk = optimized_risk * np.sqrt(52)
print("Risque Annuel du Porte-feuille(en %):", annual_risk*100)


# Tracer la frontière efficiente
plt.figure(figsize=(10, 6))
plt.plot(risks, returns, 'y-o', markersize=5, label='Frontière Efficiente')
plt.title('Frontière Efficiente selon le Modèle de Markowitz')
plt.xlabel('Risque (Écart-Type)')
plt.ylabel('Rendement Attendu')
plt.grid(True)

# Tracer des portefeuilles individuels
for i, symbol in enumerate(symbols):
    plt.scatter(np.sqrt(cov_matrix.iloc[i, i]), mean_returns[i], label=symbol)

# Tracer le portefeuille optimisé
plt.scatter(optimized_risk, optimized_return, color='red', marker='*', s=100, label='Portefeuille Optimisé')

plt.legend()
plt.show()
