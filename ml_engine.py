"""Machine-learning layer for PathPilot.

The initial model is trained on generated, rule-consistent training examples so the
application works without an external dataset. As real learner history grows, this
module can be retrained on anonymized historical outcomes.
"""
import random
import numpy as np
from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import Pipeline

FEATURE_NAMES = [
    'avg_quiz_score', 'study_minutes', 'missed_sessions',
    'days_remaining', 'topics_remaining', 'mastery_trend'
]

class PathPilotML:
    def __init__(self, seed=42):
        self.seed = seed
        self.risk_model = self._train_risk_model()
        self.time_model = self._train_time_model()

    def _make_samples(self, n=1800):
        rng = random.Random(self.seed)
        X, y_risk, y_days = [], [], []
        for _ in range(n):
            avg = rng.uniform(20, 100)
            minutes = rng.uniform(20, 180)
            missed = rng.randint(0, 12)
            remaining = rng.randint(1, 90)
            topics = rng.randint(1, 20)
            trend = rng.uniform(-25, 20)
            X.append([avg, minutes, missed, remaining, topics, trend])
            # Synthetic educational outcome used only to bootstrap a local model.
            risk_score = (60-avg)*0.9 + missed*6 - min(minutes,120)*0.12 - trend*0.7 + topics*0.8
            y_risk.append(1 if risk_score > 25 else 0)
            daily_effective = max(0.25, minutes/60) * max(0.55, min(1.15, avg/80))
            estimated = max(1, round(topics * 1.8 / daily_effective + max(0, -trend)*0.12))
            y_days.append(estimated)
        return np.array(X), np.array(y_risk), np.array(y_days)

    def _train_risk_model(self):
        X, y, _ = self._make_samples()
        return Pipeline([
            ('scale', StandardScaler()),
            ('model', RandomForestClassifier(n_estimators=180, random_state=self.seed, class_weight='balanced'))
        ]).fit(X, y)

    def _train_time_model(self):
        X, _, y = self._make_samples()
        return Pipeline([
            ('scale', StandardScaler()),
            ('model', RandomForestRegressor(n_estimators=180, random_state=self.seed))
        ]).fit(X, y)

    def predict(self, avg_quiz_score, study_minutes, missed_sessions, days_remaining, topics_remaining, mastery_trend):
        x = np.array([[avg_quiz_score, study_minutes, missed_sessions, days_remaining, topics_remaining, mastery_trend]])
        prob = float(self.risk_model.predict_proba(x)[0][1])
        label = 'High' if prob >= .67 else ('Medium' if prob >= .34 else 'Low')
        days = max(1, int(round(float(self.time_model.predict(x)[0]))))
        return {
            'risk_label': label,
            'risk_probability': round(prob * 100),
            'estimated_days': days,
            'model': 'Random Forest',
            'features': FEATURE_NAMES,
        }

ml = PathPilotML()
