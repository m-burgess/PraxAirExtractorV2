package sample;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.fxml.FXML;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception{
        Parent root = FXMLLoader.load(getClass().getResource("PraxAirExtractorV2.fxml"));
        primaryStage.setTitle("Praxair Extraction Tool");
        primaryStage.setScene(new Scene(root, 1500, 700));
        primaryStage.show();
    }


    public static void main(String[] args) {
        launch(args);
    }
}
