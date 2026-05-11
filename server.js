const http = require("node:http");


const server = http.createServer();

server.on("request", (req, rsp) => {
    console.log(`[REQUEST] ${req.method} ${req.url}`);

    console.log("header", req.rawHeaders);

    if (req.method === "GET") {
        if (req.url === "/") {
            return rsp.end("PING");
        }

        return rsp.end("NO GET HANDLER!");
    }

    if (req.method === "POST") {
        if (req.url === "/mixed-data") {
            const chunks = [];

            console.log("BEFORE data");
            req.on("data", (chunk) => {
                console.log("req data", chunk);
                chunks.push(chunk);
            });
            console.log("AFTER data");

            console.log("BEFORE req end");
            req.on("end", () => {
                const buf = Buffer.concat(chunks);
                // console.log("req end", buf.toString("ascii"));
                // console.log("req end", buf.toString("utf8"));
                // console.log("req end", buf.toString("hex"));
                // console.log("req end", buf);
                const value = buf.toString("utf8");
                const obj = JSON.parse(value);

                console.log("value", value);
                console.log("obj", obj);
                //rsp.end("REQ END");
                return rsp.end(JSON.stringify(obj));
            });

            console.log("AFTER req end");
            return;
            
        }

        return rsp.end("NO POST HANDLER!");
    }


    return rsp.end("NO END POINT!");
});
    

server.listen(8989, () => {
    console.log("Server is litenning on port ...");
});


