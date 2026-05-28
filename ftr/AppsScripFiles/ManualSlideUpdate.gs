/**
 * Refreshes all charts in a Google Slides presentation.
 * @param {string} slidesId The ID of the presentation to update.
 */
function ManuualupdateSlides() {
  console.time("updateSlides");
  try {
    const presentation = SlidesApp.openById(FTR_SLIDES_ID_PROD);
    const charts = presentation.getSlides().flatMap(slide => slide.getSheetsCharts());
    Logger.log(`Found ${charts.length} charts to refresh.`);
    charts.forEach(chart => chart.refresh());
  } catch (e) {
    Logger.log(`Error updating slides: ${e.message}`);
  }
  console.timeEnd("updateSlides");
}

