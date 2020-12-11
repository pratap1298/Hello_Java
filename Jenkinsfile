pipeline {
    agent any
    stages{
            stage('Compile')
            {
                steps
                {
                    echo "Compiled Successfully!!";
                }
            }
            stage('Junit')
            {
                steps
                {
                    echo "Junit Passsed  Successfully!!";
                }
            }
            stage('Quality-Gate')
            {
                steps
                {
                    echo "SonarQube Quality Gate passed successfully!!";
                }
            }
            stage('Deploy')
            {
                steps
                {
                    echo "PAss!";
                }
            }
        }
        post {
            always{
                echo "This will always run"
            }
            success{
                echo "This will run on;y if successful"
            }
            failure{
                echo "This will run only if failed"
            }
            unstable{
                echo "This will run only if the run was marked as unstable"
            }
            changed{
                echo "This will run only if state of pipeline has changed"
                echo " For example if the pipeline was previously faling but now it is successful"
            }
       
        }
}
