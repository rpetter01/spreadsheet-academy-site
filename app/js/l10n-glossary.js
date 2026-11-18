/* Words inside the practice spreadsheets, in Português (Brasil) and Español: headers, months, weekdays, product and category names.
   They replace the English words in the sheet data, in the answers that "Show answer" types and in what the checks accept.
   Rules: [english, português, español]; the English text must match a whole cell (or a whole "string literal" in a formula) exactly.
   Names of people (Ana, Ben...), codes (P001, SKU-...), letters, dates and numbers are left as they are.
   Lessons that match text with formulas ("West", "Pass"...) must have their lesson text translated with the same words. */
(function (root) {
'use strict';
const I = root.I18N; if (!I) return;
I.gloss([
  /* months and days */
  ['Jan', 'Jan', 'Ene'], ['Feb', 'Fev', 'Feb'], ['Mar', 'Mar', 'Mar'], ['Apr', 'Abr', 'Abr'], ['Sep', 'Set', 'Sep'],
  ['Mon', 'Seg', 'Lun'], ['Tue', 'Ter', 'Mar'], ['Wed', 'Qua', 'Mié'], ['Thu', 'Qui', 'Jue'], ['Fri', 'Sex', 'Vie'], ['Sat', 'Sáb', 'Sáb'], ['Sun', 'Dom', 'Dom'],
  ['Month', 'Mês', 'Mes'], ['Day', 'Dia', 'Día'],
  /* Foundations: headers and labels */
  ['Sales', 'Vendas', 'Ventas'], ['Revenue', 'Receita', 'Ingresos'], ['Price', 'Preço', 'Precio'], ['Quantity', 'Quantidade', 'Cantidad'], ['Qty', 'Qtd', 'Cant'], ['Units', 'Unidades', 'Unidades'],
  ['Budget', 'Orçamento', 'Presupuesto'], ['Spent', 'Gasto', 'Gastado'], ['Left', 'Restante', 'Restante'], ['Name', 'Nome', 'Nombre'], ['Product', 'Produto', 'Producto'], ['Region', 'Região', 'Región'],
  ['Item', 'Item', 'Artículo'], ['Title', 'Título', 'Título'], ['Report', 'Relatório', 'Informe'], ['Deal', 'Negócio', 'Operación'],
  ['Student', 'Aluno', 'Alumno'], ['Score', 'Nota', 'Nota'], ['Lowest', 'Menor', 'Menor'], ['Highest', 'Maior', 'Mayor'], ['Average', 'Média', 'Promedio'], ['Count', 'Contagem', 'Recuento'],
  ['Target (+10%)', 'Meta (+10%)', 'Objetivo (+10%)'], ['Tax rate', 'Alíquota', 'Tasa de impuesto'], ['Tax', 'Imposto', 'Impuesto'], ['Rate', 'Taxa', 'Tasa'], ['Fee', 'Tarifa', 'Tarifa'], ['Tip', 'Gorjeta', 'Propina'], ['Disc', 'Desc', 'Dto'],
  ['From', 'De', 'Desde'], ['To', 'Até', 'Hasta'], ['Approved', 'Aprovado', 'Aprobado'], ['Draft', 'Rascunho', 'Borrador'],
  ['Quarterly Reprot', 'Relatório trimestal', 'Informe trimestal'], ['Quarterly Report', 'Relatório trimestral', 'Informe trimestral'],
  ['Busiest day', 'Dia mais movimentado', 'Día más concurrido'], ['Best seller', 'Mais vendido', 'Más vendido'],
  ['here', 'aqui', 'aquí'], ['Jump here', 'Pule aqui', 'Salta aquí'], ['Select me', 'Selecione-me', 'Seleccióname'], ['Me too!', 'Eu também!', 'Yo también!'],
  /* things that are sold or used in examples */
  ['Pen', 'Caneta', 'Bolígrafo'], ['Book', 'Livro', 'Libro'], ['Mug', 'Caneca', 'Taza'], ['Lamp', 'Luminária', 'Lámpara'], ['Bag', 'Bolsa', 'Bolsa'], ['Cap', 'Boné', 'Gorra'], ['Tea', 'Chá', 'Té'], ['Desk', 'Mesa', 'Escritorio'],
  /* Data Skills */
  ['Result', 'Resultado', 'Resultado'], ['Rep', 'Vendedor', 'Vendedor'], ['Bonus', 'Bônus', 'Bono'], ['Age', 'Idade', 'Edad'], ['Member', 'Membro', 'Socio'], ['Entry', 'Entrada', 'Entrada'], ['Yes', 'Sim', 'Sí'], ['No', 'Não', 'No'], ['Blue', 'Azul', 'Azul'], ['Red', 'Vermelho', 'Rojo'],
  ['Order', 'Pedido', 'Pedido'], ['Shipping', 'Frete', 'Envío'], ['Target', 'Meta', 'Objetivo'], ['Actual', 'Real', 'Real'], ['Beat it?', 'Bateu a meta?', '¿Superó la meta?'], ['Working age?', 'Idade de trabalho?', '¿Edad laboral?'],
  ['Value', 'Valor', 'Valor'], ['Number?', 'Número?', '¿Número?'], ['n/a', 'n/d', 'n/d'], ['Email', 'E-mail', 'Correo'], ['Check', 'Verificação', 'Comprobación'],
  ['Tier', 'Nível', 'Nivel'], ['Label', 'Rótulo', 'Etiqueta'], ['Manager', 'Gerente', 'Gerente'], ['South', 'Sul', 'Sur'], ['Discount', 'Desconto', 'Descuento'], ['Gold', 'Ouro', 'Oro'], ['Bronze', 'Bronze', 'Bronce'], ['Silver', 'Prata', 'Plata'], ['Platinum', 'Platina', 'Platino'], ['Weight', 'Peso', 'Peso'],
  ['Phone', 'Smartphone', 'Smartphone'], ['Tablet', 'Tablet', 'Tableta'], ['West deals', 'Negócios Oeste', 'Operaciones Oeste'], ['North avg', 'Média Norte', 'Promedio Norte'], ['Over 700', 'Acima de 700', 'Más de 700'], ['Starts T', 'Começa com T', 'Empieza por T'],
  ['Dept', 'Depto', 'Dpto'], ['Status', 'Status', 'Estado'], ['Active', 'Ativo', 'Activo'], ['Support', 'Suporte', 'Soporte'], ['Ops', 'Operações', 'Operaciones'], ['Entries', 'Entradas', 'Entradas'], ['Missing', 'Ausente', 'Falta'], ['Numeric', 'Numérico', 'Numérico'],
  ['Support active', 'Suporte ativo', 'Soporte activo'], ['People', 'Pessoas', 'Personas'], ['Scored', 'Com nota', 'Con nota'], ['No score', 'Sem nota', 'Sin nota'], ['Sales active', 'Vendas ativo', 'Ventas activo'],
  ['Category', 'Categoria', 'Categoría'], ['Amount', 'Montante', 'Importe'], ['Travel', 'Viagens', 'Viajes'], ['Meals', 'Refeições', 'Comidas'], ['Sales meals', 'Refeições Vendas', 'Comidas Ventas'], ['Over 400', 'Acima de 400', 'Más de 400'],
  ['Grand total', 'Total geral', 'Total general'], ['Sales dept', 'Depto Vendas', 'Dpto Ventas'], ['Ops software', 'Software Operações', 'Software Operaciones'], ['Max travel', 'Maior viagem', 'Viaje máximo'],
  ['Code', 'Código', 'Código'], ['City', 'Cidade', 'Ciudad'], ['Number', 'Número', 'Número'], ['Messy', 'Bagunçado', 'Desordenado'], ['Clean', 'Limpo', 'Limpio'], ['First', 'Nome', 'Nombre'], ['Last', 'Sobrenome', 'Apellido'], ['Full', 'Completo', 'Completo'], ['Digits', 'Dígitos', 'Dígitos'],
  ['User', 'Usuário', 'Usuario'], ['Domain', 'Domínio', 'Dominio'], ['Position', 'Posição', 'Posición'], ['Ending', 'Final', 'Final'], ['Colour', 'Cor', 'Color'], ['Zone', 'Zona', 'Zona'], ['Num', 'Nº', 'Nº'], ['Names', 'Nomes', 'Nombres'], ['List', 'Lista', 'Lista'], ['Text', 'Texto', 'Texto'], ['Fixed', 'Fixo', 'Fijo'], ['Display', 'Exibição', 'Visualización'], ['Sentence', 'Frase', 'Frase'],
  ['Notebook', 'Caderno', 'Cuaderno'], ['Backpack', 'Mochila', 'Mochila'], ['Bottle', 'Garrafa', 'Botella'], ['Kettle', 'Chaleira', 'Hervidor'], ['Min score', 'Nota mínima', 'Nota mínima'], ['Grade', 'Conceito', 'Calificación'], ['Cost', 'Custo', 'Coste'], ['From kg', 'A partir de kg', 'Desde kg'], ['Pick', 'Escolha', 'Elección'],
  ['Not found', 'Não encontrado', 'No encontrado'], ['Not stocked', 'Sem estoque', 'Sin existencias'], ['Back*', 'Moch*', 'Moch*'],
  ['Median', 'Mediana', 'Mediana'], ['Most common', 'Mais comum', 'Más común'], ['Range', 'Amplitude', 'Rango'], ['Top 3 total', 'Total dos 3 maiores', 'Total de los 3 mayores'],
  ['Raw', 'Bruto', 'Bruto'], ['Rounded', 'Arredondado', 'Redondeado'], ['Items', 'Itens', 'Artículos'], ['Snapped', 'Ajustado', 'Ajustado'],
  ['Event', 'Evento', 'Evento'], ['Date', 'Data', 'Fecha'], ['Launch', 'Lançamento', 'Lanzamiento'], ['Holiday', 'Feriado', 'Festivo'], ['Year', 'Ano', 'Año'], ['Start', 'Início', 'Inicio'], ['End', 'Fim', 'Fin'], ['Days', 'Dias', 'Días'], ['Born', 'Nascimento', 'Nacimiento'], ['On', 'Em', 'El'],
  ['Invoice', 'Fatura', 'Factura'], ['Month end', 'Fim do mês', 'Fin de mes'], ['Today', 'Hoje', 'Hoy'], ['Days since', 'Dias desde', 'Días desde'],
  ['Store', 'Loja', 'Tienda'], ['Line total', 'Total da linha', 'Total de línea'], ['Net', 'Líquido', 'Neto'], ['Downtown', 'Centro', 'Centro'], ['Grocery', 'Mercearia', 'Alimentación'], ['Paid', 'Pago', 'Pagado'], ['Mall', 'Shopping', 'Centro comercial'],
  ['Electronics', 'Eletrônicos', 'Electrónica'], ['Toys', 'Brinquedos', 'Juguetes'], ['Refunded', 'Reembolsado', 'Reembolsado'], ['Airport', 'Aeroporto', 'Aeropuerto'], ['Min spend', 'Gasto mínimo', 'Gasto mínimo'],
  ['Net revenue', 'Receita líquida', 'Ingresos netos'], ['Refund rate', 'Taxa de reembolso', 'Tasa de reembolso'], ['Span (days)', 'Intervalo (dias)', 'Intervalo (días)'], ['Mall tier', 'Nível do shopping', 'Nivel del centro'], ['Paid electronics qty > 1', 'Eletrônicos pagos qtd > 1', 'Electrónica pagada cant > 1'],
  /* answers typed in formulas */
  ['Pass', 'Aprovado', 'Aprobado'], ['Fail', 'Reprovado', 'Reprobado'], ['Allowed', 'Permitido', 'Permitido'], ['Denied', 'Negado', 'Denegado'], ['Free', 'Grátis', 'Gratis'], ['None', 'Nenhum', 'Ninguno'], ['Unassigned', 'Sem atribuição', 'Sin asignar'],
  /* Reporting */
  ['Rent', 'Aluguel', 'Alquiler'], ['Food', 'Alimentação', 'Comida'], ['Quarter', 'Trimestre', 'Trimestre'], ['Avg food', 'Média alimentação', 'Promedio comida'], ['West total', 'Total Oeste', 'Total Oeste'],
  ['Laptop $', 'Notebook $', 'Portátil $'], ['Laptop', 'Notebook', 'Portátil'],
  /* Reporting: charts and dashboard */
  ['May', 'Mai', 'May'], ['Jun', 'Jun', 'Jun'], ['Costs', 'Custos', 'Costes'], ['Profit', 'Lucro', 'Beneficio'], ['Total sales', 'Total de vendas', 'Ventas totales'], ['Total profit', 'Lucro total', 'Beneficio total'], ['Best month', 'Melhor mês', 'Mejor mes'],
  ['Transport', 'Transporte', 'Transporte'], ['Fun', 'Lazer', 'Ocio'], ['Other', 'Outros', 'Otros'], ['Watch', 'Relógio', 'Reloj'], ['Camera', 'Câmera', 'Cámara'], ['Week', 'Semana', 'Semana'], ['Visits', 'Visitas', 'Visitas'],
  ['Week 1', 'Semana 1', 'Semana 1'], ['Week 2', 'Semana 2', 'Semana 2'], ['Week 3', 'Semana 3', 'Semana 3'], ['Week 4', 'Semana 4', 'Semana 4'], ['Week 5', 'Semana 5', 'Semana 5'], ['Week 6', 'Semana 6', 'Semana 6'], ['Week 7', 'Semana 7', 'Semana 7'], ['Week 8', 'Semana 8', 'Semana 8'],
  /* Foundations extras (formatting, rows and columns, paste special, find and replace) */
  ['Prices exclude tax', 'Preços sem imposto', 'Precios sin impuestos'], ['Notes', 'Observações', 'Notas'], ['Open', 'Aberto', 'Abierto'], ['Quarterly revenue report', 'Relatório trimestral de receita', 'Informe trimestral de ingresos'], ['Operating costs', 'Custos operacionais', 'Costes operativos'],
  ['Test', 'Teste', 'Prueba'], ['Cup', 'Xícara', 'Taza'], ['Hours', 'Horas', 'Horas'], ['Doubled', 'Dobrado', 'Doblado'], ['Full name', 'Nome completo', 'Nombre completo'], ['First name', 'Primeiro nome', 'Nombre'], ['Area code', 'DDD', 'Prefijo'],
  ['Company', 'Empresa', 'Empresa'], ['Banana', 'Banana', 'Plátano'], ['Canal', 'Canal', 'Canal'],
  /* Data Skills extras */
  ['Customers', 'Clientes', 'Clientes'], ['Each', 'Cada', 'Cada'], ['none', 'nenhum', 'ninguno'], ['Sum', 'Soma', 'Suma'], ['Amounts', 'Valores', 'Importes'], ['Copy this 1', 'Copie este 1', 'Copia este 1'], ['TBD', 'A definir', 'Por definir'],
  ['Record', 'Registro', 'Registro'], ['Ticket', 'Chamado', 'Incidencia'], ['Rating', 'Avaliação', 'Valoración'],
  ['Day no.', 'Nº do dia', 'Nº de día'], ['Priority', 'Prioridade', 'Prioridad'], ['Level', 'Nível', 'Nivel'], ['Working days', 'Dias úteis', 'Días laborables'], ['Order date', 'Data do pedido', 'Fecha del pedido'], ['Ship date', 'Data de envio', 'Fecha de envío'],
  ['Loan', 'Empréstimo', 'Préstamo'], ['Annual rate', 'Taxa anual', 'Tasa anual'], ['Months', 'Meses', 'Meses'], ['Payment', 'Parcela', 'Pago'], ['Cash flow', 'Fluxo de caixa', 'Flujo de caja'], ['NPV', 'VPL', 'VAN'], ['Total interest', 'Juros totais', 'Intereses totales'],
  ['Closed', 'Fechado', 'Cerrado'], ['Pending', 'Pendente', 'Pendiente'],
  /* chart titles the learner types */
  ['Monthly Sales 2024', 'Vendas mensais 2024', 'Ventas mensuales 2024'], ['Sales (USD)', 'Vendas (USD)', 'Ventas (USD)'], ['Sales by Region', 'Vendas por região', 'Ventas por región'], ['Sales vs Costs', 'Vendas vs Custos', 'Ventas vs Costes'], ['Monthly Performance', 'Desempenho mensal', 'Rendimiento mensual'], ['Profit by month', 'Lucro por mês', 'Beneficio por mes'],
  /* Power User track + extras */
  ['Commission','Comissão','Comisión'],['Per unit','Por unidade','Por unidad'],['Qtr','Trim.','Trim.'],['East laptops','Notebooks Leste','Portátiles Este'],['West > 500','Oeste > 500','Oeste > 500'],['West > 120','Oeste > 120','Oeste > 120'],['East > 700','Leste > 700','Este > 700'],
  ['Exam','Prova','Examen'],['Quiz','Teste','Prueba'],['Project','Projeto','Proyecto'],['Midterm','Prova parcial','Parcial'],['Final','Final','Final'],
  ['Salary','Salário','Salario'],['IT','TI','TI'],['HR','RH','RR. HH.'],['Min','Mín.','Mín.'],['West phone','Smartphone Oeste','Smartphone Oeste'],['East laptop','Notebook Leste','Portátil Este'],
  ['Regions','Regiões','Regiones'],['Products','Produtos','Productos'],['West reps','Vendedores Oeste','Vendedores Oeste'],['Ranked','Classificado','Clasificado'],
  ['Flag','Alerta','Alerta'],['Star','Estrela','Estrella'],['Stars','Estrelas','Estrellas'],['Best deal','Melhor negócio','Mejor trato'],['Best rep','Melhor vendedor','Mejor vendedor'],
  ['Area','Área','Área'],['Type','Tipo','Tipo'],['Beds','Quartos','Habitaciones'],['Sqft','ft²','ft²'],['$/sqft','$/ft²','$/ft²'],['Harbor','Baía','Bahía'],['House','Casa','Casa'],['Condo','Apartamento','Apartamento'],['Sold','Vendido','Vendido'],['Midtown','Midtown','Midtown'],['Hillside','Colina','Colina'],
  ['Avg sold house','Média casas vendidas','Promedio casas vendidas'],['Median active','Mediana ativos','Mediana activos'],['Record sale area','Área da maior venda','Área de la mayor venta'],['Hillside sold %','% vendido Colina','% vendido Colina'],['Cheapest 3+ bed','Mais barato 3+ quartos','Más barato 3+ habitaciones'],
  ['Patient','Paciente','Paciente'],['Ward','Ala','Servicio'],['Admitted','Internado','Ingresado'],['Discharged','Alta','Alta'],['Insured','Segurado','Asegurado'],['Stay (days)','Permanência (dias)','Estancia (días)'],
  ['Cardiology','Cardiologia','Cardiología'],['Surgery','Traumatologia','Traumatología'],['Pediatrics','Pediatria','Pediatría'],['Still admitted','Ainda internados','Aún ingresados'],['Cardiology avg stay','Permanência média Cardiologia','Estancia media Cardiología'],['Uninsured cost','Custo sem seguro','Coste sin seguro'],['Cardiology occupancy','Ocupação Cardiologia','Ocupación Cardiología'],['Longest stay','Maior permanência','Mayor estancia'],
  ['Item #','Item nº','Artículo nº'],['Last, First','Sobrenome, Nome','Apellido, Nombre'],['Temp','Temp.','Temp.'],['Delivery takes three to five working days after the order is confirmed','A entrega leva de três a cinco dias úteis após a confirmação do pedido','La entrega tarda de tres a cinco días hábiles tras confirmarse el pedido'],
  ['North-7', 'Norte-7', 'Norte-7'], ['East', 'Leste', 'Este'], ['West', 'Oeste', 'Oeste'], ['North', 'Norte', 'Norte'], ['Cookie', 'Biscoito', 'Galleta']
]);
})(window);
