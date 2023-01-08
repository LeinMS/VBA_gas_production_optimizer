# VBA_gas_production_optimizer

Перед созданием скрипта стоял основной вопрос: как с одной стороны моделировать эксплуатацию месторождения, чтобы поддерживать постоянную загрузку на 
ЦПС/УКПГ и при этом не превышать показатели. 

Принцип работы:

Были построены зависимости, с помощью которых рассчитываются множители на штуцер.
- В случае превышения добычи газа, мы выбираем скважины в зависимости от их добычи. 
  - Если скважина добывает много газа, то она сильнее зажимается. 
  - Если скважина добывает мало газа, то сильнее разжимается.

Сначала мы делали оптимизацию на каждом временном шаге (провели временной шаг – посмотрели на результат – применили на следующий шаг). 
После этого модифицировали оптимизатор и делаем в одном временном шаге несколько промежуточных шагов. Если не удовлетворяет условию, меняем штуцера,
пересчитываем и достигаем сходимости. 
Простыми словами, держим добычу в некотором коридоре в районе 29500 – 30500 млн. кубических метров.
Когда коридор достигается, это становится сигналом для перехода на следующий временной шаг.
