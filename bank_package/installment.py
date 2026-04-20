class InstallmentCalculator:
    """Расчет рассрочки"""
    
    def __init__(self, amount, months):
        self.amount = amount
        self.months = months
    
    def monthly_payment(self):
        """Ежемесячный платеж"""
        return round(self.amount / self.months, 2)
    
    def total_payment(self):
        """Общая сумма"""
        return self.amount