FROM rust:1.86-slim AS builder

WORKDIR /app

COPY . .

RUN cargo build --release -p ch03_common

FROM debian:bookworm-slim

WORKDIR /app

COPY --from=builder /app/target/release/ch03_common /usr/local/bin/ch03_common

CMD ["/usr/local/bin/ch03_common"]