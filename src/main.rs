use std::collections::HashMap;
use std::process::Command;
use regex::Regex;

fn main() {
    match get_windows_version() {
        Ok(version) => {
            println!("检测到系统版本: {}", version);
            if let Some(kms_key) = get_kms_key(&version) {
                println!("正在使用密钥激活: {}", kms_key);
                activate_windows(&kms_key);
            } else {
                eprintln!("未找到匹配的KMS密钥，请手动指定");
            }
        }
        Err(e) => eprintln!("版本检测失败: {}", e),
    }
    let _ = Command::new("cmd").arg("/c").arg("pause").status();
}

// KMS密钥数据库（根据微软文档更新）
fn get_kms_keys() -> HashMap<&'static str, &'static str> {
    let mut keys = HashMap::new();

    // Windows 11
    keys.insert("Windows 11 Pro", "W269N-WFGWX-YVC9B-4J6C9-T83GX");
    keys.insert("Windows 11 Enterprise", "NPPR9-FWDCX-D2C8J-H872K-2YT43");
    keys.insert("Windows 11 Education", "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2");

    // Windows 10
    keys.insert("Windows 10 Pro", "W269N-WFGWX-YVC9B-4J6C9-T83GX");
    keys.insert("Windows 10 Enterprise", "NPPR9-FWDCX-D2C8J-H872K-2YT43");
    keys.insert("Windows 10 Education", "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2");
    keys.insert("Windows 10 LTSC 2021", "M7XTQ-FN8P6-TTKYV-9D4CC-J462D");

    // Windows Server
    keys.insert("Windows Server 2022", "WX4NM-KYWYW-QJJR4-XV3QB-6VM33");
    keys.insert("Windows Server 2019", "WMDGN-G9PQG-XVVXX-R3X43-63DFG");
    keys.insert("Windows Server 2016", "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY");
    keys.insert("Windows Server 2012 R2", "D2N9P-3P6X9-2R39C-7RTCD-MDVJX");
    keys.insert("Windows Server 2012", "BN3D2-R7TKB-3YPBD-8DRP2-27GG4");
    keys.insert("Windows Server 2008 R2", "YC6KT-GKW9T-YTKYR-T4X34-R7VHC");

    // 其他版本
    keys.insert("Windows 8.1 Pro", "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9");
    keys.insert("Windows 7 Professional", "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4");

    keys
}

fn get_kms_key(version: &str) -> Option<&'static str> {
    let keys = get_kms_keys();

    // 精确匹配优先
    if let Some(key) = keys.get(version) {
        return Some(key);
    }

    // 模糊匹配逻辑
    let patterns = [
        (r"(?i)Windows 11 Pro", "Windows 11 Pro"),
        (r"(?i)Windows 11", "Windows 11 Pro"),         // 默认使用Pro版密钥
        (r"(?i)Windows 10 LTSC", "Windows 10 LTSC 2021"),
        (r"(?i)Windows 10", "Windows 10 Pro"),        // 默认使用Pro版密钥
        (r"(?i)Windows Server 2022", "Windows Server 2022"),
        (r"(?i)Windows Server 2019", "Windows Server 2019"),
        (r"(?i)Windows Server 2016", "Windows Server 2016"),
        (r"(?i)Windows 8\.1", "Windows 8.1 Pro"),
        (r"(?i)Windows 7", "Windows 7 Professional"),
    ];

    for (pattern, key) in patterns {
        if Regex::new(pattern).unwrap().is_match(version) {
            return keys.get(key).copied();
        }
    }

    None
}

fn activate_windows(kms_key: &str) {
    let commands = [
        format!("slmgr /ipk {}", kms_key),
        "slmgr /skms 192.168.253.169".to_string(),
        "slmgr /ato".to_string(),
    ];

    for cmd in commands {
        println!("执行命令: {}", cmd);
        let output = Command::new("cmd")
            .args(&["/C", &cmd])
            .output()
            .expect("命令执行失败");

        println!("{}", String::from_utf8_lossy(&output.stdout));
        if !output.status.success() {
            eprintln!("错误: {}", String::from_utf8_lossy(&output.stderr));
        }
    }
}

// 之前的版本检测函数（保持原有实现）
fn get_windows_version() -> Result<String, Box<dyn std::error::Error>> {
    let output = Command::new("cmd")
        .args(&["/C", "systeminfo"])
        .output()?;

    let system_info = String::from_utf8_lossy(&output.stdout);

    let re = Regex::new(r"(?im)^OS\s*[^:\r\n]*:\s*Microsoft\s*(Windows[^\r\n]*)")?;

    if let Some(caps) = re.captures(&system_info) {
        let full_version = caps[1].trim();

        let cleaned = full_version
            .replace(|c: char| !c.is_ascii_alphanumeric() && c != ' ' && c != '.' && c != '®' && c != '™', "")
            .replace("®", "")
            .replace("™", "")
            .replace(" 专业版", "").replace(" Pro", "")
            .replace(" 家庭版", "").replace(" Home", "")
            .replace(" 企业版", "").replace(" Enterprise", "")
            .replace(" 教育版", "").replace(" Education", "")
            .replace(" 数据中心版", "").replace(" Datacenter", "")
            .replace(" 标准版", "").replace(" Standard", "")
            .replace(" 核心版", "").replace(" Core", "")
            .replace(" 中文版", "")
            .replace(" China", "")
            .replace(" for Workstations", "")
            .replace("  ", " ")
            .trim()
            .to_string();

        let version_re = Regex::new(
            r"(?i)(Windows\s(?:Server\s)?(?:10|11|\d{4}|XP|Vista|7|8\.?1?|2000|2003|2008\s?R2|2012\s?R2|201[6-9]|202\d))"
        )?;

        if let Some(caps) = version_re.captures(&cleaned) {
            return Ok(caps[1].to_string());
        }

        return Ok(cleaned);
    }

    Err("无法检测系统版本".into())
}