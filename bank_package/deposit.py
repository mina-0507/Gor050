class DepositCalculator:
    """Расчет вклада"""
    
    def __init__(self, amount, rate, months):
        self.amount = amount
        self.rate = rate / 100 / 12
        self.months = months
    
    def final_amount(self):
        """Итоговая сумма"""
        return round(self.amount * (1 + self.rate) ** self.months, 2)
    
    def total_interest(self):
        """Начисленные проценты"""
        return round(self.final_amount() - self.amount, 2)
    
    def get_schedule(self):
        """График начисления процентов по месяцам"""
        schedule = []
        amount = self.amount
        
        for month in range(1, self.months + 1):
            interest = amount * self.rate
            amount = amount + interest
            
            schedule.append({
                'month': month,
                'amount': round(amount, 2),
                'interest': round(interest, 2)
            })
        
        return schedule