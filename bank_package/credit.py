class CreditCalculator:
    """Расчет кредита"""
    
    def __init__(self, amount, rate, months):
        self.amount = amount
        self.rate = rate / 100 / 12
        self.months = months
    
    def monthly_payment(self):
        """Ежемесячный платеж"""
        if self.rate == 0:
            return self.amount / self.months
        k = (self.rate * (1 + self.rate) ** self.months) / ((1 + self.rate) ** self.months - 1)
        return round(self.amount * k, 2)
    
    def total_payment(self):
        """Общая сумма выплат"""
        return round(self.monthly_payment() * self.months, 2)
    
    def overpayment(self):
        """Переплата"""
        return round(self.total_payment() - self.amount, 2)
    
    def get_schedule_annuity(self):
        """График платежей по месяцам"""
        schedule = []
        remaining = self.amount
        payment = self.monthly_payment()
        
        for month in range(1, self.months + 1):
            interest = remaining * self.rate
            principal = payment - interest
            remaining = remaining - principal
            
            schedule.append({
                'month': month,
                'payment': round(payment, 2),
                'principal': round(principal, 2),
                'interest': round(interest, 2),
                'remaining': round(max(remaining, 0), 2)
            })
        
        return schedule