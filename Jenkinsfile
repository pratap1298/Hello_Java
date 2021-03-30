#!/usr/bin/env groovy

// Jenkinsfile version 0.1
// declarative pipeline

pipeline {
    agent any

   

    stages {

        stage("Stage 1") {
            steps {
                script {
                    properties([parameters([file(description: 'Upload HSDP Cost excel sheet( ex : imcs-billing-report-yyyy-mm-dd.xlsx)', name: 'imcs-billing-report'), 
                    file(description: 'Upload HSDP Trend excel sheet(ex:HSDP-Trend.xlsx)', name: 'HSDP-Trend')])])
                    echo "HSDP Cost excel file path is ${params.imcs-billing-report}"
                    echo "HSDP Trend  excel file path is ${params.HSDP-Trend}"

                    

                }
            }
        }

        stage("stage 2") {
            steps {
                script {
                    println("Stage 2") 
                    sh "HSDP-Trend.py ${params.imcs-billing-report} ${params.HSDP-Trend} " 
                }
            }
        }

    } //end of stages
} //end of pipeline


