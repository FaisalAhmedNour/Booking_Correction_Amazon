// Faisal Ahmed (M) - 11 August 2025

import Box from "@mui/material/Box";
import Button from "@mui/material/Button";
import IconButton from "@mui/material/IconButton";
import Modal from "@mui/material/Modal";
import FormTitle from "../../Components/FormTitle";
import FormControl from "@mui/material/FormControl";
import MenuItem from "@mui/material/MenuItem";
import Select from "@mui/material/Select";
import TextField from "@mui/material/TextField";
import Checkbox from "@mui/material/Checkbox";
import FormControlLabel from "@mui/material/FormControlLabel";
import UploadFileOutlinedIcon from "@mui/icons-material/UploadFileOutlined";
import CloseIcon from '@mui/icons-material/Close';
import { getStoredVendorOptions, getStoredCompanyOptions } from "./lib";
import { useMemo } from "react";

const style = {
    position: 'absolute',
    top: '50%',
    left: '50%',
    transform: 'translate(-50%, -50%)',
    width: 620,
    maxHeight: '92vh',
    overflowY: 'auto',
    bgcolor: 'background.paper',
    border: '2px solid #000',
    boxShadow: 24,
    p: 2,
    borderRadius: 2,
    borderColor: '#1976d2',
};

const inputLength = 200;

const ProcessController = ({
    isProcessing,
    isStartVisible,
    engineError,
    handleStop,
    handleStart,
    isGettingData,
    setIsGettingData,
    open,
    handleClose,
    formData,
    setFormData
}) => {
    const currentVendorOptions = useMemo(() => getStoredVendorOptions(), [open]);
    const currentCompanyOptions = useMemo(() => getStoredCompanyOptions(), [open]);

    const getExcelFilePath = async () => {
        setIsGettingData(true);
        const filePath = await window?.engine?.openFile({
            title: "Select File",
            accept: ["openFile"],
        });
        if (filePath?.success === true) {
            setFormData(prev => ({ ...prev, excelFile: filePath?.path?.[0] }))
            setIsGettingData(false);
        }
    }

    const getFolderPath = async () => {
        setIsGettingData(true);
        const filePath = await window?.engine?.openFile({
            title: "Select Folder",
            accept: ["openDirectory"],
        });
        if (filePath?.success === true) {
            setFormData(prev => ({ ...prev, downloadDir: filePath?.path?.[0] }))
            setIsGettingData(false);
        }
    }

    const handleCompanyChange = (e) => {
        const value = e.target.value;
        setFormData(prev => ({
            ...prev,
            settings: {
                ...prev?.settings,
                companyCode: value
            }
        }));
    };

    const handleVendorChange = (e) => {
        const value = e.target.value;
        const selectedVendor = currentVendorOptions.find(v => v.value === value || v.vendorFullName === value || v.shipVendorCode === value);
        setFormData(prev => ({
            ...prev,
            settings: {
                ...prev?.settings,
                shipVendorCode: value,
                addEditManufacturer: selectedVendor ? (selectedVendor.addEditManufacturer ?? "") : (prev?.settings?.addEditManufacturer || ""),
                shipperName: selectedVendor ? (selectedVendor.shipperName ?? "") : (prev?.settings?.shipperName || "Epic Garments DWC LLC"),
                contactPhone: selectedVendor ? (selectedVendor.contactPhone ?? "") : (prev?.settings?.contactPhone || ""),
                contactEmail: selectedVendor ? (selectedVendor.contactEmail ?? "") : (prev?.settings?.contactEmail || ""),
                emailAddressToReceiveFCR: selectedVendor ? (selectedVendor.emailAddressToReceiveFCR ?? "") : (prev?.settings?.emailAddressToReceiveFCR || "")
            }
        }));
    };

    const handleSettingChange = (field, value) => {
        setFormData(prev => ({
            ...prev,
            settings: {
                ...prev?.settings,
                [field]: value
            }
        }));
    };

    return (
        <Modal open={open} onClose={handleClose}>
            <Box sx={style}>
                <div className="relative z-20 px-1">
                    <span className="absolute -top-4 -right-4">
                        <IconButton size="small" color="error" onClick={handleClose} >
                            <CloseIcon />
                        </IconButton>
                    </span>
                    <form
                        onSubmit={handleStart}
                        className="flex flex-col items-center gap-1.5 justify-center bg-white p-2"
                    >
                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Username"}
                                isCompulsory={true}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <TextField
                                    placeholder="Enter username"
                                    size="small"
                                    name="userName"
                                    variant="outlined"
                                    value={formData?.userName ?? ""}
                                    type="text"
                                    sx={{
                                        width: "100%",
                                        "& .MuiInputBase-root": {
                                            height: 25,
                                            fontSize: 14,
                                        },
                                        overflow: "hidden",
                                    }}
                                    className="editableInput"
                                    onChange={(e) => setFormData(prev => ({ ...prev, userName: e.target.value }))}
                                />
                            </div>
                        </div>

                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Password"}
                                isCompulsory={true}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <TextField
                                    placeholder="Enter password"
                                    size="small"
                                    name="password"
                                    variant="outlined"
                                    value={formData?.password ?? ""}
                                    type="password"
                                    sx={{
                                        width: "100%",
                                        "& .MuiInputBase-root": {
                                            height: 25,
                                            fontSize: 14,
                                        },
                                        overflow: "hidden",
                                    }}
                                    className="editableInput"
                                    onChange={(e) => setFormData(prev => ({ ...prev, password: e.target.value }))}
                                />
                            </div>
                        </div>

                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Select Excel File"}
                                isCompulsory={true}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <TextField
                                    inputProps={{ readOnly: true }}
                                    placeholder="Select file"
                                    size="small"
                                    name="excelFile"
                                    variant="outlined"
                                    value={formData?.excelFile ?? ""}
                                    sx={{
                                        width: "100%",
                                        bgcolor: "#e8f0fe",
                                        "& .MuiInputBase-root": {
                                            height: 25,
                                            fontSize: 14,
                                        },
                                        overflow: "hidden",
                                    }}
                                    className="fixedInput"
                                />
                                <Button
                                    disabled={isGettingData}
                                    sx={{
                                        position: "absolute",
                                        right: "1px",
                                        height: "92%",
                                        width: "6px",
                                        marginTop: "1px",
                                        bgcolor: "white",
                                        color: "#283e8a",
                                        ":hover": {
                                            color: "white",
                                        },
                                    }}
                                    component="label"
                                    variant="contained"
                                    onClick={getExcelFilePath}
                                >
                                    <UploadFileOutlinedIcon />
                                </Button>
                            </div>
                        </div>

                        {/* Settings: Company Code Dropdown */}
                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Company Code"}
                                isCompulsory={true}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <FormControl size="small" sx={{ width: "100%" }}>
                                    <Select
                                        value={formData?.settings?.companyCode ?? ""}
                                        onChange={handleCompanyChange}
                                        displayEmpty
                                        sx={{
                                            height: 25,
                                            fontSize: 13,
                                            "& .MuiSelect-select": {
                                                py: 0,
                                                fontSize: 13,
                                                display: "flex",
                                                alignItems: "center"
                                            }
                                        }}
                                    >
                                        {currentCompanyOptions.map((item, idx) => (
                                            <MenuItem key={idx} value={item.companyCode} sx={{ fontSize: 13 }}>
                                                {item.companyCode}
                                            </MenuItem>
                                        ))}
                                    </Select>
                                </FormControl>
                            </div>
                        </div>

                        {/* Settings: Ship/Vendor Code Dropdown */}
                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Ship/Vendor Code"}
                                isCompulsory={true}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <FormControl size="small" sx={{ width: "100%" }}>
                                    <Select
                                        value={formData?.settings?.shipVendorCode ?? ""}
                                        onChange={handleVendorChange}
                                        displayEmpty
                                        renderValue={(selected) => {
                                            if (!selected) {
                                                return <span className="text-gray-400">Select Vendor / Factory</span>;
                                            }
                                            const found = currentVendorOptions.find(v => v.value === selected || v.vendorFullName === selected);
                                            return found ? (found.factory ? `${found.factory} (${found.shipVendorCode})` : found.shipVendorCode) : selected;
                                        }}
                                        sx={{
                                            height: 25,
                                            fontSize: 13,
                                            "& .MuiSelect-select": {
                                                py: 0,
                                                fontSize: 13,
                                                display: "flex",
                                                alignItems: "center"
                                            }
                                        }}
                                    >
                                        {currentVendorOptions.map((item, idx) => (
                                            <MenuItem key={idx} value={item.value} sx={{ fontSize: 13 }}>
                                                {item.factory ? `${item.factory} (${item.shipVendorCode})` : item.shipVendorCode}
                                            </MenuItem>
                                        ))}
                                    </Select>
                                </FormControl>
                            </div>
                        </div>

                        {/* Settings: Add/Edit Manufacturer */}
                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Add/Edit Manufacturer"}
                                isCompulsory={false}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <TextField
                                    placeholder="Enter manufacturer"
                                    size="small"
                                    name="addEditManufacturer"
                                    variant="outlined"
                                    value={formData?.settings?.addEditManufacturer ?? ""}
                                    type="text"
                                    sx={{
                                        width: "100%",
                                        "& .MuiInputBase-root": {
                                            height: 25,
                                            fontSize: 14,
                                        },
                                        overflow: "hidden",
                                    }}
                                    className="editableInput"
                                    onChange={(e) => handleSettingChange("addEditManufacturer", e.target.value)}
                                />
                            </div>
                        </div>

                        {/* Settings: Shipper Name */}
                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Shipper Name"}
                                isCompulsory={false}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <TextField
                                    placeholder="Enter shipper name"
                                    size="small"
                                    name="shipperName"
                                    variant="outlined"
                                    value={formData?.settings?.shipperName ?? ""}
                                    type="text"
                                    sx={{
                                        width: "100%",
                                        "& .MuiInputBase-root": {
                                            height: 25,
                                            fontSize: 14,
                                        },
                                        overflow: "hidden",
                                    }}
                                    className="editableInput"
                                    onChange={(e) => handleSettingChange("shipperName", e.target.value)}
                                />
                            </div>
                        </div>

                        {/* Settings: Contact Phone */}
                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Contact Phone (optional)"}
                                isCompulsory={false}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <TextField
                                    placeholder="Enter contact phone"
                                    size="small"
                                    name="contactPhone"
                                    variant="outlined"
                                    value={formData?.settings?.contactPhone ?? ""}
                                    type="text"
                                    sx={{
                                        width: "100%",
                                        "& .MuiInputBase-root": {
                                            height: 25,
                                            fontSize: 14,
                                        },
                                        overflow: "hidden",
                                    }}
                                    className="editableInput"
                                    onChange={(e) => handleSettingChange("contactPhone", e.target.value)}
                                />
                            </div>
                        </div>

                        {/* Settings: Contact Email */}
                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Contact Email"}
                                isCompulsory={false}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <TextField
                                    placeholder="Enter contact email"
                                    size="small"
                                    name="contactEmail"
                                    variant="outlined"
                                    value={formData?.settings?.contactEmail ?? ""}
                                    type="email"
                                    sx={{
                                        width: "100%",
                                        "& .MuiInputBase-root": {
                                            height: 25,
                                            fontSize: 14,
                                        },
                                        overflow: "hidden",
                                    }}
                                    className="editableInput"
                                    onChange={(e) => handleSettingChange("contactEmail", e.target.value)}
                                />
                            </div>
                        </div>

                        {/* Settings: Email Address to Receive FCR */}
                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Email to receive FCR"}
                                isCompulsory={false}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <TextField
                                    placeholder="Enter email to receive FCR"
                                    size="small"
                                    name="emailAddressToReceiveFCR"
                                    variant="outlined"
                                    value={formData?.settings?.emailAddressToReceiveFCR ?? ""}
                                    type="email"
                                    sx={{
                                        width: "100%",
                                        "& .MuiInputBase-root": {
                                            height: 25,
                                            fontSize: 14,
                                        },
                                        overflow: "hidden",
                                    }}
                                    className="editableInput"
                                    onChange={(e) => handleSettingChange("emailAddressToReceiveFCR", e.target.value)}
                                />
                            </div>
                        </div>

                        {/* Settings: Status Dropdown */}
                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Status"}
                                isCompulsory={true}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <FormControl size="small" sx={{ width: "100%" }}>
                                    <Select
                                        value={formData?.settings?.status ?? "Draft"}
                                        onChange={(e) => handleSettingChange("status", e.target.value)}
                                        displayEmpty
                                        sx={{
                                            height: 25,
                                            fontSize: 13,
                                            "& .MuiSelect-select": {
                                                py: 0,
                                                fontSize: 13,
                                                display: "flex",
                                                alignItems: "center"
                                            }
                                        }}
                                    >
                                        <MenuItem value="Draft" sx={{ fontSize: 13 }}>Draft</MenuItem>
                                        <MenuItem value="Sent" sx={{ fontSize: 13 }}>Sent</MenuItem>
                                    </Select>
                                </FormControl>
                            </div>
                        </div>

                        {/* Settings: Test Mode Dropdown */}
                        <div className="flex justify-start items-center gap-1">
                            <FormTitle
                                text={"Test Mode"}
                                isCompulsory={true}
                                length={inputLength}
                            />
                            <div className="flex-grow mx-auto relative w-[320px]">
                                <FormControl size="small" sx={{ width: "100%" }}>
                                    <Select
                                        value={formData?.settings?.testMode ?? (formData?.settings?.reviewMode ? "Yes" : "No")}
                                        onChange={(e) => {
                                            const val = e.target.value;
                                            setFormData(prev => ({
                                                ...prev,
                                                settings: {
                                                    ...prev?.settings,
                                                    testMode: val,
                                                    reviewMode: val === "Yes"
                                                }
                                            }));
                                        }}
                                        displayEmpty
                                        sx={{
                                            height: 25,
                                            fontSize: 13,
                                            "& .MuiSelect-select": {
                                                py: 0,
                                                fontSize: 13,
                                                display: "flex",
                                                alignItems: "center"
                                            }
                                        }}
                                    >
                                        <MenuItem value="No" sx={{ fontSize: 13 }}>No</MenuItem>
                                        <MenuItem value="Yes" sx={{ fontSize: 13 }}>Yes</MenuItem>
                                    </Select>
                                </FormControl>
                            </div>
                        </div>

                        {typeof engineError === 'string' &&
                            <p className="text-[red] text-xs text-center mt-1 col-span-2 order-11">{engineError}</p>}
                        <div className="flex items-center gap-2 mt-2">
                            <Button
                                type="submit"
                                variant="outlined"
                                size="small"
                                sx={{ height: 25 }}
                                color="success"
                            >
                                Start Correction
                            </Button>
                        </div>
                    </form>
                </div>
            </Box>
        </Modal>
    )
}

export default ProcessController;