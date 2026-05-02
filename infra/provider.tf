terraform {
  required_version = "~> 1.9"

  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 5.0"
    }
  }

  # Backend is configured dynamically via -backend-config flags in CI
  # Run bootstrap-terraform-backend.yml workflow first to create these resources
  backend "s3" {}
}

provider "aws" {
  region = var.aws_region

  default_tags {
    tags = {
      Project   = "shu-chatbot"
      ManagedBy = "terraform"
    }
  }
}
