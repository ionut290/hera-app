package it.vargacantieri.hera.camera;

import android.app.Activity;
import android.content.Intent;
import android.graphics.Color;
import android.graphics.Typeface;
import android.graphics.drawable.GradientDrawable;
import android.os.Bundle;
import android.view.Gravity;
import android.view.View;
import android.view.ViewGroup;
import android.widget.Button;
import android.widget.FrameLayout;
import android.widget.TextView;

import androidx.annotation.NonNull;
import androidx.camera.core.CameraSelector;
import androidx.camera.core.ImageCapture;
import androidx.camera.core.ImageCaptureException;
import androidx.camera.core.Preview;
import androidx.camera.lifecycle.ProcessCameraProvider;
import androidx.camera.view.PreviewView;
import androidx.core.content.ContextCompat;
import androidx.fragment.app.FragmentActivity;

import com.google.common.util.concurrent.ListenableFuture;

import java.io.File;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.UUID;

public class HeraContinuousCameraActivity extends FragmentActivity {
    public static final String EXTRA_MAX_PHOTOS = "maxPhotos";
    public static final String EXTRA_PHOTO_PATHS = "photoPaths";

    private static final int DEFAULT_MAX_PHOTOS = 10;
    private static final int MIN_CONTROL_SIZE_DP = 56;

    private final ArrayList<String> photoPaths = new ArrayList<>();
    private ImageCapture imageCapture;
    private TextView counterView;
    private Button shutterButton;
    private Button doneButton;
    private File sessionFolder;
    private int maxPhotos = DEFAULT_MAX_PHOTOS;
    private boolean captureInProgress = false;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        getWindow().setStatusBarColor(Color.BLACK);
        getWindow().setNavigationBarColor(Color.BLACK);

        maxPhotos = Math.max(1, Math.min(getIntent().getIntExtra(EXTRA_MAX_PHOTOS, DEFAULT_MAX_PHOTOS), DEFAULT_MAX_PHOTOS));
        sessionFolder = new File(getCacheDir(), "hera-continuous-camera/" + UUID.randomUUID());
        if (!sessionFolder.mkdirs() && !sessionFolder.isDirectory()) {
            finishWithError();
            return;
        }

        PreviewView previewView = new PreviewView(this);
        previewView.setScaleType(PreviewView.ScaleType.FILL_CENTER);

        FrameLayout root = new FrameLayout(this);
        root.setBackgroundColor(Color.BLACK);
        root.addView(previewView, new FrameLayout.LayoutParams(
            ViewGroup.LayoutParams.MATCH_PARENT,
            ViewGroup.LayoutParams.MATCH_PARENT
        ));

        counterView = makeCounterView();
        FrameLayout.LayoutParams counterParams = new FrameLayout.LayoutParams(
            ViewGroup.LayoutParams.WRAP_CONTENT,
            dp(44)
        );
        counterParams.gravity = Gravity.TOP | Gravity.START;
        counterParams.setMargins(dp(18), dp(22), 0, 0);
        root.addView(counterView, counterParams);

        doneButton = makeDoneButton();
        FrameLayout.LayoutParams doneParams = new FrameLayout.LayoutParams(
            dp(92),
            dp(44)
        );
        doneParams.gravity = Gravity.TOP | Gravity.END;
        doneParams.setMargins(0, dp(22), dp(18), 0);
        root.addView(doneButton, doneParams);

        shutterButton = makeShutterButton();
        FrameLayout.LayoutParams shutterParams = new FrameLayout.LayoutParams(dp(82), dp(82));
        shutterParams.gravity = Gravity.BOTTOM | Gravity.CENTER_HORIZONTAL;
        shutterParams.setMargins(0, 0, 0, dp(34));
        root.addView(shutterButton, shutterParams);

        TextView hint = makeHintView();
        FrameLayout.LayoutParams hintParams = new FrameLayout.LayoutParams(
            ViewGroup.LayoutParams.WRAP_CONTENT,
            ViewGroup.LayoutParams.WRAP_CONTENT
        );
        hintParams.gravity = Gravity.BOTTOM | Gravity.CENTER_HORIZONTAL;
        hintParams.setMargins(dp(16), 0, dp(16), dp(126));
        root.addView(hint, hintParams);

        setContentView(root);
        updateCounter();
        doneButton.setOnClickListener(view -> finishWithPhotos());
        shutterButton.setOnClickListener(view -> takePhoto());
        startCamera(previewView);
    }

    @Override
    public void onBackPressed() {
        cleanupSession();
        setResult(Activity.RESULT_CANCELED);
        super.onBackPressed();
    }

    private void startCamera(PreviewView previewView) {
        ListenableFuture<ProcessCameraProvider> future = ProcessCameraProvider.getInstance(this);
        future.addListener(() -> {
            try {
                ProcessCameraProvider provider = future.get();
                Preview preview = new Preview.Builder().build();
                preview.setSurfaceProvider(previewView.getSurfaceProvider());

                imageCapture = new ImageCapture.Builder()
                    .setCaptureMode(ImageCapture.CAPTURE_MODE_MINIMIZE_LATENCY)
                    .setJpegQuality(85)
                    .build();

                provider.unbindAll();
                provider.bindToLifecycle(
                    this,
                    CameraSelector.DEFAULT_BACK_CAMERA,
                    preview,
                    imageCapture
                );
            } catch (Exception error) {
                finishWithError();
            }
        }, ContextCompat.getMainExecutor(this));
    }

    private void takePhoto() {
        if (imageCapture == null || captureInProgress || photoPaths.size() >= maxPhotos) return;

        captureInProgress = true;
        shutterButton.setEnabled(false);
        int photoNumber = photoPaths.size() + 1;
        String fileName = "Foto-" + new DecimalFormat("00").format(photoNumber) + ".jpg";
        File file = new File(sessionFolder, fileName);

        ImageCapture.OutputFileOptions options = new ImageCapture.OutputFileOptions.Builder(file).build();
        imageCapture.takePicture(
            options,
            ContextCompat.getMainExecutor(this),
            new ImageCapture.OnImageSavedCallback() {
                @Override
                public void onImageSaved(@NonNull ImageCapture.OutputFileResults outputFileResults) {
                    photoPaths.add(file.getAbsolutePath());
                    captureInProgress = false;
                    updateCounter();
                    if (photoPaths.size() < maxPhotos) {
                        shutterButton.setEnabled(true);
                    }
                }

                @Override
                public void onError(@NonNull ImageCaptureException exception) {
                    if (file.exists()) file.delete();
                    captureInProgress = false;
                    shutterButton.setEnabled(photoPaths.size() < maxPhotos);
                    counterView.setText("Errore · riprova");
                }
            }
        );
    }

    private void updateCounter() {
        int count = photoPaths.size();
        counterView.setText(count == 1 ? "1 foto" : count + " foto");
        if (count >= maxPhotos) {
            counterView.setText(count + " foto · limite");
            shutterButton.setEnabled(false);
        }
        doneButton.setEnabled(count > 0 && !captureInProgress);
        doneButton.setAlpha(doneButton.isEnabled() ? 1f : 0.55f);
    }

    private void finishWithPhotos() {
        if (captureInProgress || photoPaths.isEmpty()) return;
        Intent result = new Intent();
        result.putStringArrayListExtra(EXTRA_PHOTO_PATHS, new ArrayList<>(photoPaths));
        setResult(Activity.RESULT_OK, result);
        finish();
    }

    private void finishWithError() {
        cleanupSession();
        setResult(Activity.RESULT_CANCELED);
        finish();
    }

    private void cleanupSession() {
        if (sessionFolder == null || !sessionFolder.exists()) return;
        File[] files = sessionFolder.listFiles();
        if (files != null) {
            for (File file : files) file.delete();
        }
        sessionFolder.delete();
    }

    private TextView makeCounterView() {
        TextView view = new TextView(this);
        view.setTextColor(Color.WHITE);
        view.setTextSize(16);
        view.setTypeface(Typeface.DEFAULT, Typeface.BOLD);
        view.setGravity(Gravity.CENTER);
        view.setPadding(dp(14), 0, dp(14), 0);
        view.setBackground(roundedBackground(0xA6000000, 22));
        return view;
    }

    private Button makeDoneButton() {
        Button button = new Button(this);
        button.setAllCaps(false);
        button.setText("Fine");
        button.setTextSize(16);
        button.setTypeface(Typeface.DEFAULT, Typeface.BOLD);
        button.setTextColor(Color.WHITE);
        button.setPadding(0, 0, 0, 0);
        button.setMinHeight(dp(MIN_CONTROL_SIZE_DP));
        button.setMinWidth(dp(MIN_CONTROL_SIZE_DP));
        button.setBackground(roundedBackground(0xCC087A5B, 22));
        return button;
    }

    private Button makeShutterButton() {
        Button button = new Button(this);
        button.setText("");
        button.setMinWidth(0);
        button.setMinHeight(0);
        GradientDrawable outer = new GradientDrawable();
        outer.setShape(GradientDrawable.OVAL);
        outer.setColor(Color.WHITE);
        outer.setStroke(dp(6), 0x99FFFFFF);
        button.setBackground(outer);
        button.setElevation(dp(8));
        return button;
    }

    private TextView makeHintView() {
        TextView view = new TextView(this);
        view.setText("Scatta tutte le foto · poi premi Fine");
        view.setTextColor(Color.WHITE);
        view.setTextSize(14);
        view.setGravity(Gravity.CENTER);
        view.setPadding(dp(14), dp(8), dp(14), dp(8));
        view.setBackground(roundedBackground(0x99000000, 18));
        return view;
    }

    private GradientDrawable roundedBackground(int color, int radiusDp) {
        GradientDrawable background = new GradientDrawable();
        background.setColor(color);
        background.setCornerRadius(dp(radiusDp));
        return background;
    }

    private int dp(int value) {
        return Math.round(value * getResources().getDisplayMetrics().density);
    }
}
