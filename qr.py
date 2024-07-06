import stripe

stripe.api_key = 'your_stripe_secret_key'

def generate_payment_link(amount, currency='inr'):
    session = stripe.checkout.Session.create(
        payment_method_types=['card'],
        line_items=[{
            'price_data': {
                'currency': currency,
                'product_data': {
                    'name': 'Milk Bill Payment',
                },
                'unit_amount': amount * 100,
            },
            'quantity': 1,
        }],
        mode='payment',
        success_url='https://yourwebsite.com/success?session_id={CHECKOUT_SESSION_ID}',
        cancel_url='https://yourwebsite.com/cancel',
    )
    return session.url
