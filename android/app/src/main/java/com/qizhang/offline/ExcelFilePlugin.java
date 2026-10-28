package com.qizhang.offline;

import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.util.Base64;

import androidx.activity.result.ActivityResult;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.ActivityCallback;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.io.OutputStream;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ConcurrentHashMap;
import java.util.Map;

@CapacitorPlugin(name = "ExcelFile")
public class ExcelFilePlugin extends Plugin {
    private final ExecutorService executor = Executors.newSingleThreadExecutor();
    private final Map<String, byte[]> pendingData = new ConcurrentHashMap<>();

    @PluginMethod
    public void save(PluginCall call) {
        String fileName = call.getString("fileName", "Qizhang.xlsx");
        String mimeType = call.getString("mimeType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        String base64 = call.getString("data");
        if (base64 == null || base64.isEmpty()) {
            call.reject("没有可保存的 Excel 数据");
            return;
        }

        Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        intent.setType(mimeType);
        intent.putExtra(Intent.EXTRA_TITLE, fileName);
        try {
            pendingData.put(call.getCallbackId(), Base64.decode(base64, Base64.DEFAULT));
        } catch (IllegalArgumentException error) {
            call.reject("Excel 数据格式无效", error);
            return;
        }
        startActivityForResult(call, intent, "saveFileResult");
    }

    @ActivityCallback
    private void saveFileResult(PluginCall call, ActivityResult result) {
        if (call == null) return;
        byte[] bytes = pendingData.remove(call.getCallbackId());
        if (result.getResultCode() != Activity.RESULT_OK || result.getData() == null || result.getData().getData() == null) {
            call.reject("已取消保存", "SAVE_CANCELLED");
            return;
        }
        Uri uri = result.getData().getData();
        if (bytes == null || bytes.length == 0) {
            call.reject("Excel 数据已失效，请重新导出");
            return;
        }
        executor.execute(() -> {
            try (OutputStream output = getContext().getContentResolver().openOutputStream(uri, "wt")) {
                if (output == null) throw new IllegalStateException("无法写入所选位置");
                output.write(bytes);
                output.flush();
                JSObject response = new JSObject();
                response.put("saved", true);
                response.put("bytesWritten", bytes.length);
                call.resolve(response);
            } catch (Exception error) {
                call.reject("保存 Excel 文件失败", error);
            }
        });
    }
}
