import * as React from "react";
import Box from "@mui/material/Box";
import TextField from "@mui/material/TextField";
import { Button } from "@mui/material";

export default function InputForm(props) {
    const [key, setKey] = React.useState("");
    const [val, setVal] = React.useState("");
    const addKeyValue = props.addKeyValue;
    return (
        <Box
            component="form"
            sx={{
                "& > :not(style)": { m: 1, width: "20ch" }
            }}
            noValidate
            autoComplete="off"
        >
            <div
                style={{
                    display: "flex",
                    flexDirection: "column",
                    justifyContent: "space-between",
                    width: "300px"
                }}
            >
                <TextField id="outlined-basic" label="please enter key" value={key} onChange={(e) => { setKey(e.target.value) }} />
                <br/>
                <TextField id="filled-basic" label="please enter value" value={val} onChange={(e) => { setVal(e.target.value) }} />
            </div>

            <Button onClick={() => {
                console.log(key, val);
                addKeyValue({ "name": key, "value": val });
                setKey("");
                setVal("");
            }} size="small" variant="contained">
                Add Property
            </Button>

            {/* <TextField id="standard-basic"  label="Standard" variant="standard" /> */}
        </Box>
    );
}