export_model_table_labels <- function(model, file_path) {
  # Load required packages
  require(broom)
  require(dplyr)
  require(flextable)
  require(officer)
  require(stringr)
  
  # Extract model results
  results <- broom::tidy(model, conf.int = TRUE)
  
  # Clean up variable names
  results <- results %>%
    mutate(
      Variable = term %>%
        str_replace_all(":", " * ") %>%
        str_replace_all("\\(Intercept\\)", "Intercept") %>%
        str_replace_all("Gender1", "Gender[Female]") %>%
        str_replace_all("GroupADHD\\-Inattentive Type", "ADHD[IN]") %>%
        str_replace_all("GroupADHD\\-Combined Type", "ADHD[Comb]") %>%
        str_replace_all("age", "Age") %>%
        str_replace_all("iaf", "IAF")
    )
  
  # Custom sorting logic on the cleaned Variable labels
  results <- results %>%
    mutate(
      sort_order = case_when(
        # Intercept
        Variable == "Intercept" ~ 0,
        
        # ADHD[IN] terms — always ordered before ADHD[Comb]
        Variable == "ADHD[IN]" ~ 101,
        Variable == "ADHD[IN] * Age" ~ 102,
        Variable == "ADHD[IN] * Gender[Female]" ~ 103,
        Variable == "ADHD[IN] * IAF" ~ 104,
        Variable == "ADHD[IN] * Age * Gender[Female]" ~ 105,
        Variable == "ADHD[IN] * Age * IAF" ~ 106,
        Variable == "ADHD[IN] * Gender[Female] * IAF" ~ 107,
        Variable == "ADHD[IN] * Age * Gender[Female] * IAF" ~ 108,
        
        # ADHD[Comb] terms — always after ADHD[IN]
        Variable == "ADHD[Comb]" ~ 201,
        Variable == "ADHD[Comb] * Age" ~ 202,
        Variable == "ADHD[Comb] * Gender[Female]" ~ 203,
        Variable == "ADHD[Comb] * IAF" ~ 204,
        Variable == "ADHD[Comb] * Age * Gender[Female]" ~ 205,
        Variable == "ADHD[Comb] * Age * IAF" ~ 206,
        Variable == "ADHD[Comb] * Gender[Female] * IAF" ~ 207,
        Variable == "ADHD[Comb] * Age * Gender[Female] * IAF" ~ 208,
        
        # Main (non-group) effects and interactions
        Variable %in% c(
          "Age", "Gender[Female]", "IAF",
          "Age * Gender[Female]", "Age * IAF", "Gender[Female] * IAF",
          "Age * Gender[Female] * IAF"
        ) ~ 50 + as.numeric(factor(Variable, levels = c(
          "Age", "Gender[Female]", "IAF",
          "Age * Gender[Female]", "Age * IAF", "Gender[Female] * IAF",
          "Age * Gender[Female] * IAF"
        ))),
        
        # Fallback
        TRUE ~ 999
      )
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
    arrange(sort_order) %>%
    select(Variable, β, SE, CI, `t-value`, `p-value`)
  
  # Create flextable
  ft <- flextable(results) %>% autofit()
  
  # Export to Word
  doc <- read_docx() %>%
    body_add_flextable(ft)
  
  print(doc, target = file_path)
}
