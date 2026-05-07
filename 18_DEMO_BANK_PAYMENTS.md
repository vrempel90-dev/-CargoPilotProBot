# Demo Bank Payments

Добавлена виртуальная демо-оплата через варианты:

- Kaspi
- Halyk
- ЦентрКредит / BCC
- Freedom

Важно:
это НЕ реальная оплата и НЕ списывает деньги. Это имитация банковской оплаты для демонстрации клиенту.

Сценарий:
1. клиент открывает `/track/CG...`;
2. нажимает `💳 Оплатить`;
3. выбирает Kaspi / Halyk / ЦентрКредит / Freedom;
4. видит демо-страницу выбранного банка;
5. нажимает `Оплатить в демо`;
6. заказ отмечается как оплаченный;
7. на странице груза меняется статус оплаты.

Страницы:
- `/demo-pay/CG...` — выбор банка
- `/demo-pay/CG.../checkout/kaspi`
- `/demo-pay/CG.../checkout/halyk`
- `/demo-pay/CG.../checkout/bcc`
- `/demo-pay/CG.../checkout/freedom`

API:
- `/api/demo-payment/create/CG...?provider=kaspi`
- `/api/demo-payment/create/kaspi/CG...`
- `/api/demo-payment/success/CG...?provider=kaspi&payment_id=...`
- `/api/demo-payment/success/kaspi/CG...?payment_id=...`

В реальном внедрении вместо демо-страниц подключается официальный эквайринг банка или платёжного сервиса.
