/* "Skip step": the button, the confirmation, the tag and the end-of-lesson note, in Português and Español. */
(function (root) {
'use strict';
const I = root.I18N; if (!I) return;
I.ui([
  ['Skip step', 'Pular etapa', 'Saltar paso'], ['Skip this step', 'Pular esta etapa', 'Saltar este paso'], ['Skipped', 'Pulada', 'Omitido'],
  ['Skip this step? You will not earn XP for it, and you can come back to it at any time.', 'Pular esta etapa? Você não ganha XP por ela e pode voltar a ela quando quiser.', '¿Saltar este paso? No ganarás XP por él y puedes volver a él cuando quieras.'],
  ['You skipped 1 step. You can come back to it at any time.', 'Você pulou 1 etapa. Pode voltar a ela quando quiser.', 'Te saltaste 1 paso. Puedes volver a él cuando quieras.']
]);
I.ui([
  ['To fill down: select the cell with the formula and the cells below it, then press <b>Ctrl + D</b>. With only one cell selected, Ctrl + D copies the cell above it.', 'Para preencher para baixo: selecione a célula com a fórmula e as células abaixo dela, depois pressione <b>Ctrl + D</b>. Com uma única célula selecionada, o Ctrl + D copia a célula de cima.', 'Para rellenar hacia abajo: selecciona la celda con la fórmula y las celdas de abajo, y pulsa <b>Ctrl + D</b>. Con una sola celda seleccionada, Ctrl + D copia la celda de arriba.']
]);
I.ui([
  ['To fill down: select <b>{r}</b> and press <b>Ctrl + D</b>. With only one cell selected, Ctrl + D copies the cell above it.', 'Para preencher para baixo: selecione <b>{r}</b> e pressione <b>Ctrl + D</b>. Com uma única célula selecionada, o Ctrl + D copia a célula de cima.', 'Para rellenar hacia abajo: selecciona <b>{r}</b> y pulsa <b>Ctrl + D</b>. Con una sola celda seleccionada, Ctrl + D copia la celda de arriba.']
]);
I.pat(/^You skipped (\d+) steps\. You can come back to them at any time\.$/, 'Você pulou $1 etapas. Pode voltar a elas quando quiser.', 'Te saltaste $1 pasos. Puedes volver a ellos cuando quieras.');
})(window);
