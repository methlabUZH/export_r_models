export_model_table <- function(model, file_path) {
  # Load required packages
  require(broom)
  require(dplyr)
  require(flextable)
  require(officer)
  require(stringr)
  
  # Extract model results
  results <- broom::tidy(model, conf.int = TRUE)
  
  # Prettify term names
  results <- results %>%
    mutate(
      Variable = term %>%
        str_replace_all(":", " * ") %>%
        str_replace_all("\\((Intercept)\\)", "Intercept") %>%
        str_replace_all("([A-Za-z_]+)([A-Z][a-z0-9]+$)", "\\1 [\\2]")
    )
  
  # Format statistics
  results <- results %>%
    mutate(
      β = round(estimate, 2),
      SE = round(std.error, 2),
      CI = paste0(round(conf.low, 2), " – ", round(conf.high, 2)),
      `t-value` = round(statistic, 2),
      p_string = ifelse(
        p.value <= 0.001,
        paste0("p = ", formatC(p.value, format = "e", digits = 2)),
        paste0("p = ", formatC(p.value, format = "f", digits = 3))
      ),
      p_stars = case_when(
        p.value < 0.001 ~ "***",
        p.value < 0.01  ~ "**",
        p.value < 0.05  ~ "*",
        TRUE            ~ ""
      ),
      `p-value` = paste0(p_string, p_stars)
    ) %>%
    select(Variable, β, SE, CI, `t-value`, `p-value`)
  
  # Create flextable
  ft <- flextable(results) %>% autofit()
  
  # Export to Word
  doc <- read_docx() %>%
    body_add_flextable(ft)
  
  print(doc, target = file_path)
}
