#!/usr/bin/env groovy

// Jenkinsfile version 0.1
// declarative pipeline

pipeline {
    agent any

   

    stages {

        stage("Stage 1") {
            steps {
                script {
                    properties([parameters([file(description: 'Upload HSDP Cost excel sheet( ex : imcs-billing-report-yyyy-mm-dd.xlsx)', name: 'C:\\WINDOWS\\system32\\config\\systemprofile\\AppData\\Local\\Jenkins\\.jenkins\\workspace\\pipeline\\imcs-billing-report-2021-03-14.xlsx'), 
                                            file(description: 'Upload HSDP Cost excel sheet( ex : imcs-billing-report-yyyy-mm-dd.xlsx)', name: 'C:\\WINDOWS\\system32\\config\\systemprofile\\AppData\\Local\\Jenkins\\.jenkins\\workspace\\pipeline\\HSDP-Trend.xlsx')])])
                    
                    echo "HSDP Trend  excel file path is updated"

                    

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


