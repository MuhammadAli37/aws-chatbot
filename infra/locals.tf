locals {
  name_prefix = "${var.app_name}-${var.environment}"

  common_tags = {
    Project     = var.app_name
    Environment = var.environment
    ManagedBy   = "terraform"
  }
}
