<script>
    let count = 0;

    function increaseCounter() {
        count += 1;
    }

    function decreaseCounter() {
        if(count !== 0) {
            count -= 1;
        }
    }

</script>

<style>

    #background {
        height: 100vh;
        width: 100vw;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    #counter {
        height: 95vmin;
        width: 70vmin;
        background: #C8C8C8;
        border-radius: 8%;
        display: flex;
        flex-direction: column;
    }

    #screen-background {
        background: #303030;
        height: 40vmin;
        width: 50vmin;
        align-self: center;
        border-radius: 5%;
        margin-top: 5vmin;
        display: flex;
        justify-content: center;
        align-items: center;
    }

    #screen {
        background: #DAF7A6;
        width: 27vmin;
        height: 27vmin;
        display: flex;
        justify-content: center;
        align-items: center;

    }

    #increase-btn {
        border-radius: 100%;
        height: 12vmin;
        width: 12vmin;
        font-size: 50px;
        margin-left: 50vmin;
        margin-top: 9vmin;
        background: #980000 ;
    }

    #decrease-btn {
        border-radius: 100%;
        height: 12vmin;
        width: 12vmin;
        font-size: 50px;
        margin-left: 35vmin;
        background: #980000 ;
    }

    #screen-text {
        font-size: 50px;
    }




</style>

<div id="background">
    <div id="counter">
        <div id="screen-background">
            <div id="screen">
                <p id="screen-text"> {count} </p>
            </div>
        </div>
        <button on:click={increaseCounter} id="increase-btn">+</button>
        <button on:click={decreaseCounter} id="decrease-btn">-</button>
    </div>
</div>