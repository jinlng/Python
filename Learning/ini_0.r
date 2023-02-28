# Load the ggplot2 library
library(ggplot2)

# Generate some sample data
set.seed(123)
x <- rnorm(1000, mean = 0, sd = 1)
y <- rnorm(1000, mean = 0, sd = 1)

# Histogram using base graphics
hist(x, main = "Histogram of X", xlab = "X values", ylab = "Frequency")

# Scatter plot using ggplot2
data <- data.frame(x = x, y = y)
ggplot(data, aes(x = x, y = y)) +
  geom_point() +
  labs(title = "Scatter plot of X and Y", x = "X values", y = "Y values")

# Bar chart using ggplot2
set.seed(456)
data2 <- data.frame(category = sample(c("A", "B", "C"), 1000, replace = TRUE))
ggplot(data2, aes(x = category)) +
  geom_bar() +
  labs(title = "Bar chart of categories", x = "Category", y = "Count")
